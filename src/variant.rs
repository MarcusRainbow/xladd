#![allow(non_snake_case, non_camel_case_types, non_upper_case_globals)]

use std::{fmt, mem, slice};
//#[cfg(feature = "try_from")]
use crate::entrypoint::excel_free;
use crate::xlcall::xloper12;
use crate::xlcall::{
    xlbitDLLFree, xlbitXLFree, xlerrDiv0, xlerrGettingData, xlerrNA, xlerrName, xlerrNull,
    xlerrNum, xlerrRef, xlerrValue, xloper12__bindgen_ty_1, xloper12__bindgen_ty_1__bindgen_ty_3,
    xltypeBool, xltypeErr, xltypeInt, xltypeMissing, xltypeMulti, xltypeNil, xltypeNum, xltypeStr,
    LPXLOPER12, XLOPER12,
};
use std::convert::TryFrom;
use std::f64;

#[derive(Debug)]
pub enum XLAddError {
    F64ConversionFailed(String),
    BoolConversionFailed(String),
    IntConversionFailed(String),
    StringConversionFailed(String),
    MissingArgument(String, String),
}

impl std::error::Error for XLAddError {}
impl fmt::Display for XLAddError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            XLAddError::F64ConversionFailed(v) => {
                write!(f, "Coud not convert parameter [{}] to f64", v)
            }
            XLAddError::BoolConversionFailed(v) => {
                write!(f, "Coud not convert parameter [{}] to bool", v)
            }
            XLAddError::IntConversionFailed(v) => {
                write!(f, "Coud not convert parameter [{}] to integer", v)
            }
            XLAddError::StringConversionFailed(v) => {
                write!(f, "Coud not convert parameter [{}] to string", v)
            }
            XLAddError::MissingArgument(func, v) => {
                write!(f, "Function [{}] is missing parameter {} ", func, v)
            }
        }
    }
}

const xltypeMask: u32 = !(xlbitDLLFree | xlbitXLFree);
const xltypeStr_xlbitDLLFree: u32 = xltypeStr | xlbitDLLFree;
const xltypeMulti_xlbitDLLFree: u32 = xltypeMulti | xlbitDLLFree;

/// Variant is a wrapper around an XLOPER12. It can contain a string, i32 or f64, or a
/// two dimensional of any mixture of these. Basically, it can contain anything that an
/// Excel cell or array of cells can contain.
pub struct Variant(XLOPER12);

impl Variant {
    /// Construct a variant containing a missing entry. This is used in function calls to
    /// signal that a parameter should be defaulted.
    pub fn missing() -> Variant {
        Variant(XLOPER12 {
            xltype: xltypeMissing,
            val: xloper12__bindgen_ty_1 { w: 0 },
        })
    }

    pub fn is_missing_or_null(&self) -> bool {
        self.0.xltype & xltypeMask == xltypeMissing || self.0.xltype & xltypeMask == xltypeNil
    }

    /// Construct a variant containing an error. This is used in Excel to represent standard errors
    /// that are shown as #DIV0 etc. Currently supported error codes are:
    /// xlerrNull, xlerrDiv0, xlerrValue, xlerrRef, xlerrName, xlerrNum, xlerrNA, xlerrGettingData
    pub fn from_err(xlerr: u32) -> Variant {
        Variant(XLOPER12 {
            xltype: xltypeErr,
            val: xloper12__bindgen_ty_1 { err: xlerr as i32 },
        })
    }

    /// Construct a variant containing an array from a slice of other variants. The variants
    /// may contain arrays or scalar strings or numbers, which are treated like single-cell
    /// arrays. They are glued either horizontally (horiz=true) or vertically. If the arrays
    /// do not match sizes in the other dimension, they are padded with blanks.
    pub fn concat(from: &[Variant], horiz: bool) -> Variant {
        // first find the size of the resulting array
        let mut columns: usize = 0;
        let mut rows: usize = 0;
        for xloper in from.iter() {
            let dim = xloper.dim();
            if horiz {
                columns += dim.0;
                rows = rows.max(dim.1);
            } else {
                columns = columns.max(dim.0);
                rows += dim.1;
            }
        }

        // Zero-sized arrays cause Excel to crash. Arrays with a dimension of
        // one (either rows or cols) are confusing to Excel, which repeats them
        // when using array paste. Solve both problems by padding with NA and
        // setting the min rows or cols to two.
        rows = rows.max(2);
        columns = columns.max(2);

        // If the array is too big, return an error string
        if rows > 1_048_576 || columns > 16384 {
            return Self::from("#ERR resulting array is too big");
        }

        // now clone the components into place
        let size = rows * columns;
        let mut array = vec![Variant::from_err(xlerrNA); size];
        let mut col = 0;
        let mut row = 0;
        for var in from.iter() {
            match var.0.xltype & xltypeMask {
                xltypeMulti => unsafe {
                    let p = var.0.val.array.lparray;
                    let var_cols = var.0.val.array.columns as usize;
                    let var_rows = var.0.val.array.rows as usize;
                    for x in 0..var_cols {
                        for y in 0..var_rows {
                            let src = (y * var_cols + x) as isize;
                            let dest = (row + y) * columns + col + x;
                            array[dest] = Variant::from(p.offset(src)).clone();
                        }
                    }

                    if horiz {
                        col += var_cols;
                    } else {
                        row += var_rows;
                    }
                },
                xltypeMissing => {}
                _ => {
                    let dest = row * columns + col;
                    array[dest] = var.clone();
                    if horiz {
                        col += 1;
                    } else {
                        row += 1;
                    }
                }
            }
        }

        let lparray = array.as_mut_ptr() as LPXLOPER12;
        mem::forget(array);

        Variant(XLOPER12 {
            xltype: xltypeMulti,
            val: xloper12__bindgen_ty_1 {
                array: xloper12__bindgen_ty_1__bindgen_ty_3 {
                    lparray,
                    rows: rows as i32,
                    columns: columns as i32,
                },
            },
        })
    }

    /// Creates a transposed clone of this Variant. If this Variant is a scalar type,
    /// simply returns it unchanged.
    pub fn transpose(&self) -> Variant {
        // simply clone any scalar type, including errors
        if (self.0.xltype & xltypeMask) != xltypeMulti {
            return self.clone();
        }

        // We have an array that we need to transpose. Create a vector of
        // Variant to contain the elements.
        let dim = self.dim();
        if dim.0 > 1_048_576 || dim.1 > 16384 {
            return Self::from("#ERR resulting array is too big");
        }

        let len = dim.0 * dim.1;
        let mut array = Vec::with_capacity(len);

        // Copy the elements transposed, cloning each one
        for i in 0..dim.1 {
            for j in 0..dim.0 {
                array.push(self.at(j, i));
            }
        }

        // Return as a Variant
        let lparray = array.as_mut_ptr() as LPXLOPER12;
        mem::forget(array);

        Variant(XLOPER12 {
            xltype: xltypeMulti,
            val: xloper12__bindgen_ty_1 {
                array: xloper12__bindgen_ty_1__bindgen_ty_3 {
                    lparray,
                    rows: dim.0 as i32,
                    columns: dim.1 as i32,
                },
            },
        })
    }

    /// Exposes the underlying XLOPER12
    pub fn as_mut_xloper(&mut self) -> &mut XLOPER12 {
        &mut self.0
    }

    /// Gets the count of rows and columns. Scalars are treated as 1x1. Missing values are
    /// treated as 0x0.
    pub fn dim(&self) -> (usize, usize) {
        match self.0.xltype & xltypeMask {
            xltypeMulti => unsafe {
                (
                    self.0.val.array.columns as usize,
                    self.0.val.array.rows as usize,
                )
            },
            xltypeMissing => (0, 0),
            _ => (1, 1),
        }
    }

    /// Gets the element at the given column and row. If this is a scalar, treat it as a one-element
    /// array. If the column or row is out of bounds, return NA. The returned element is always cloned
    /// so it can be returned as a value
    pub fn at(&self, column: usize, row: usize) -> Variant {
        if (self.0.xltype & xltypeMask) != xltypeMulti {
            if column == 0 && row == 0 {
                self.clone()
            } else {
                Self::from_err(xlerrNA)
            }
        } else {
            let (columns, rows) = unsafe {
                (
                    self.0.val.array.columns as usize,
                    self.0.val.array.rows as usize,
                )
            };
            if column >= columns || row >= rows {
                Self::from_err(xlerrNA)
            } else {
                let index = row * columns + column;
                Self::from(unsafe { self.0.val.array.lparray.add(index) }).clone()
            }
        }
    }
}

/// Construct a variant containing nil. This is used in Excel to represent cells that have
/// nothing in them. It is also a sensible starting state for an uninitialized variant.
impl Default for Variant {
    fn default() -> Variant {
        Variant(XLOPER12 {
            xltype: xltypeNil,
            val: xloper12__bindgen_ty_1 { w: 0 },
        })
    }
}

/// Converts this variant to an int. If we do not contain an int, return None. Note that
/// Excel cells do not ever contain ints, so this would only come from a non-Excel user
/// creating an XLOPER, for example the result of a call into Excel.
impl<'a> TryFrom<&'a Variant> for i32 {
    type Error = XLAddError;
    fn try_from(v: &'a Variant) -> Result<i32, Self::Error> {
        if (v.0.xltype & xltypeMask) != xltypeInt {
            Err(XLAddError::IntConversionFailed(v.to_string()))
        } else {
            Ok(unsafe { v.0.val.w })
        }
    }
}

/// Converts this variant to a float. If we do not contain a float, return Err.
impl<'a> TryFrom<&'a Variant> for f64 {
    type Error = XLAddError;
    fn try_from(v: &'a Variant) -> Result<f64, Self::Error> {
        if (v.0.xltype & xltypeMask) != xltypeNum {
            Err(XLAddError::F64ConversionFailed(v.to_string()))
        } else {
            Ok(unsafe { v.0.val.num })
        }
    }
}

/// Converts this variant to a bool. If we do not contain a bool, return Err.
impl<'a> TryFrom<&'a Variant> for bool {
    type Error = XLAddError;
    fn try_from(v: &'a Variant) -> Result<Self, Self::Error> {
        if (v.0.xltype & xltypeMask) != xltypeBool {
            Err(XLAddError::BoolConversionFailed(v.to_string()))
        } else {
            Ok(!unsafe { v.0.val.xbool == 0 })
        }
    }
}

/// Converts this variant to a string. Alternatively, you can use Display or to_string,
/// which both go through this call if the variant contains a string. Guaranteed to return
/// Some(...) if this object is of type xltypeStr. Always returns None if this object is
/// of any other type. If the string contains a unicode string that is misformed, return
/// the error message.
impl<'a> TryFrom<&'a Variant> for String {
    type Error = XLAddError;
    fn try_from(v: &'a Variant) -> Result<Self, Self::Error> {
        if (v.0.xltype & xltypeMask) != xltypeStr {
            Err(XLAddError::StringConversionFailed(v.to_string()))
        } else {
            let cstr_slice = unsafe {
                let cstr: *const u16 = v.0.val.str;
                let cstr_len = *cstr.offset(0) as usize;
                slice::from_raw_parts(cstr.offset(1), cstr_len)
            };
            match String::from_utf16(cstr_slice) {
                Ok(s) => Ok(s),
                Err(e) => Ok(e.to_string()),
            }
        }
    }
}

/// Converts a variant into a f64 array filling the missing or invalid with f64::NAN.
/// This is so that you can handle those appropriately for your application (for example fill with the mean value or 0)
impl<'a> From<&'a Variant> for Vec<f64> {
    fn from(v: &'a Variant) -> Vec<f64> {
        let (x, y) = v.dim();
        let mut res = Vec::with_capacity(x * y);
        if x == 1 && y == 1 {
            res.push(unsafe { v.0.val.num });
        } else {
            res.resize(x * y, 0.0);
            for j in 0..y {
                for i in 0..x {
                    let index = j * x + i;
                    let slice =
                        unsafe { slice::from_raw_parts::<xloper12>(v.0.val.array.lparray, x * y) };
                    let v = slice[index];
                    res[index] = if v.xltype & xltypeMask != xltypeNum {
                        0.0f64
                    } else {
                        unsafe { v.val.num }
                    };
                }
            }
        }
        res
    }
}

/// Converts a variant into a f64 array filling the missing or invalid with f64::NAN.
/// This is so that you can handle those appropriately for your application (for example fill with the mean value or 0)
impl<'a> From<&'a Variant> for Vec<String> {
    fn from(v: &'a Variant) -> Vec<String> {
        let (x, y) = v.dim();
        let mut res = Vec::with_capacity(x * y);
        if x == 1 && y == 1 {
            let cstr_slice = unsafe {
                let cstr: *const u16 = v.0.val.str;
                let cstr_len = *cstr.offset(0) as usize;
                slice::from_raw_parts(cstr.offset(1), cstr_len)
            };
            let s = match String::from_utf16(cstr_slice) {
                Ok(s) => s,
                Err(e) => e.to_string(),
            };
            res.push(s);
        } else {
            res.resize(x * y, String::new());
            for j in 0..y {
                for i in 0..x {
                    let index = j * x + i;
                    let slice =
                        unsafe { slice::from_raw_parts::<xloper12>(v.0.val.array.lparray, x * y) };
                    let v = slice[index];
                    res[index] = if (v.xltype & xltypeMask) != xltypeStr {
                        String::new()
                    } else {
                        let cstr_slice = unsafe {
                            let cstr: *const u16 = v.val.str;
                            let cstr_len = *cstr.offset(0) as usize;
                            slice::from_raw_parts(cstr.offset(1), cstr_len)
                        };
                        match String::from_utf16(cstr_slice) {
                            Ok(s) => s,
                            Err(e) => e.to_string(),
                        }
                    }
                }
            }
        }
        res
    }
}

/// Construct a variant from an LPXLOPER12, for example supplied by Excel. The assumption
/// is that Excel continues to own the XLOPER12 and its lifetime is greater than that of
/// the Variant we construct here. For example, the LPXLOPER may be an argument to one
/// of our functions. We therefore do not want to own any of the data in this variant, so
/// we clear all ownership bits. This means we treat it as a kind of dynamic mut ref.
impl From<LPXLOPER12> for Variant {
    fn from(xloper: LPXLOPER12) -> Variant {
        let mut result = Variant(unsafe { *xloper });
        result.0.xltype &= xltypeMask; // no ownership bits
        result
    }
}

/// Construct a LPXlOPER12 from a Variant. This is just a cast to the underlying union
/// contained within a pointer that we pass back to Excel. Excel will clean up the pointer
/// after us
impl From<Variant> for LPXLOPER12 {
    fn from(v: Variant) -> LPXLOPER12 {
        Box::into_raw(Box::new(v)) as LPXLOPER12
    }
}

/// Construct a variant containing an float (f64)
impl From<f64> for Variant {
    fn from(num: f64) -> Variant {
        Variant(XLOPER12 {
            xltype: xltypeNum,
            val: xloper12__bindgen_ty_1 { num },
        })
    }
}

/// Construct a variant containing an bool (i32)
impl From<bool> for Variant {
    fn from(xbool: bool) -> Variant {
        Variant(XLOPER12 {
            xltype: xltypeBool,
            val: xloper12__bindgen_ty_1 {
                xbool: xbool as i32,
            },
        })
    }
}

/// Construct a variant containing an int (i32)
impl From<i32> for Variant {
    fn from(w: i32) -> Variant {
        Variant(XLOPER12 {
            xltype: xltypeInt,
            val: xloper12__bindgen_ty_1 { w },
        })
    }
}

/// Construct a variant containing an array of strings
/// Pass in a tuple of (array, columns), it will calculate the number of rows.
impl From<&(&[&str], usize)> for Variant {
    fn from(arr: &(&[&str], usize)) -> Variant {
        let mut array = arr.0.iter().map(|&v| Variant::from(v)).collect::<Vec<_>>();
        let lparray = array.as_mut_ptr() as LPXLOPER12;
        mem::forget(array);
        let rows = if arr.1 == 0 { 0 } else { arr.0.len() / arr.1 };
        let columns = arr.1;
        if rows == 0 || columns == 0 {
            Variant::from_err(xlerrNull)
        } else if rows > 32767 {
            Variant::from_err(xlerrValue)
        } else {
            Variant(XLOPER12 {
                xltype: xltypeMulti,
                val: xloper12__bindgen_ty_1 {
                    array: xloper12__bindgen_ty_1__bindgen_ty_3 {
                        lparray,
                        rows: rows as i32,
                        columns: columns as i32,
                    },
                },
            })
        }
    }
}

/// Construct a variant containing an array of strings
/// Pass in a tuple of (array, columns), it will calculate the number of rows.
impl From<&(&[f64], usize)> for Variant {
    fn from(arr: &(&[f64], usize)) -> Variant {
        // Return as a Variant
        let mut array = arr
            .0
            .iter()
            .map(|&v| match v {
                v if v.is_nan() => Variant::from_err(xlerrNA),
                v => Variant::from(v),
            })
            .collect::<Vec<_>>();
        let rows = if arr.1 == 0 { 0 } else { arr.0.len() / arr.1 };
        let columns = arr.1;
        if rows == 0 || columns == 0 {
            Variant::from_err(xlerrNull)
        } else if rows > 32767 {
            Variant::from_err(xlerrValue)
        } else {
            let lparray = array.as_mut_ptr() as LPXLOPER12;
            mem::forget(array);
            Variant(XLOPER12 {
                xltype: xltypeMulti,
                val: xloper12__bindgen_ty_1 {
                    array: xloper12__bindgen_ty_1__bindgen_ty_3 {
                        lparray,
                        rows: rows as i32,
                        columns: columns as i32,
                    },
                },
            })
        }
    }
}

/// Construct a variant containing a string. Strings in Excel (at least after Excel 97) are 16bit
/// Unicode starting with a 16-bit length. The length is treated as signed, which means that
/// strings can be no longer than 32k characters. If a string longer than this is supplied, or a
/// string that is not valid 16bit Unicode, an xlerrValue error is stored instead.
impl From<&str> for Variant {
    fn from(s: &str) -> Variant {
        let mut wstr: Vec<u16> = s.encode_utf16().collect();
        let len = wstr.len();
        if len > 32767 {
            return Variant::from_err(xlerrValue);
        }

        // Pascal-style string with length at the start. Forget the string so we do not delete it.
        // We are now relying on the drop method of Variant to clean it up for us. Note that the
        // shrink_to_fit is essential, so the capacity is the same as the length. We have no way
        // of storing the capacity otherwise.
        wstr.insert(0, len as u16);
        wstr.shrink_to_fit();
        let p = wstr.as_mut_ptr();
        mem::forget(wstr);
        Variant(XLOPER12 {
            xltype: xltypeStr + xlbitDLLFree,
            val: xloper12__bindgen_ty_1 { str: p },
        })
    }
}

/// Construct a variant containing a string. Strings in Excel (at least after Excel 97) are 16bit
/// Unicode starting with a 16-bit length. The length is treated as signed, which means that
/// strings can be no longer than 32k characters. If a string longer than this is supplied, or a
/// string that is not valid 16bit Unicode, an xlerrValue error is stored instead.
impl From<String> for Variant {
    fn from(s: String) -> Variant {
        From::<&str>::from(s.as_str())
    }
}

/// Implement Display, which means we do not need a method for converting to strings. Just use
/// to_string.
impl fmt::Display for Variant {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match self.0.xltype & xltypeMask {
            xltypeErr => match unsafe { self.0.val.err } as u32 {
                xlerrNull => write!(f, "#NULL"),
                xlerrDiv0 => write!(f, "#DIV0"),
                xlerrValue => write!(f, "#VALUE"),
                xlerrRef => write!(f, "#REF"),
                xlerrName => write!(f, "#NAME"),
                xlerrNum => write!(f, "#NUM"),
                xlerrNA => write!(f, "#NA"),
                xlerrGettingData => write!(f, "#DATA"),
                _ => write!(f, "#BAD_ERR"),
            },
            xltypeInt => write!(f, "{}", unsafe { self.0.val.w }),
            xltypeMissing => write!(f, "#MISSING"),
            xltypeMulti => write!(f, "#MULTI"),
            xltypeNil => write!(f, "#NIL"),
            xltypeNum => write!(f, "{}", unsafe { self.0.val.num }),
            xltypeStr => write!(f, "{}", String::try_from(&self.clone()).unwrap()),
            _ => write!(f, "#BAD_XLOPER"),
        }
    }
}

/// We need to implement Drop, as Variant is a wrapper around a union type that does
/// not know how to handle its contained pointers.
impl Drop for Variant {
    fn drop(&mut self) {
        if (self.0.xltype & xlbitXLFree) != 0 {
            excel_free(&mut self.0);
            return;
        }

        match self.0.xltype {
            xltypeStr_xlbitDLLFree => {
                // We have a 16bit string that was originally allocated as a vector
                // but then forgotten. Reconstruct the vector, so its drop method
                // will clean up the memory for us.
                unsafe {
                    let p = self.0.val.str;
                    let len = *p as usize + 1;
                    let cap = len;
                    Vec::from_raw_parts(p, len, cap);
                }
            }
            xltypeMulti_xlbitDLLFree => {
                // We have an array that was originally allocated as a vector of
                // Variant but then forgotten. Reconstruct the vector, so its drop method
                // will clean up the vector and its elements for us.
                unsafe {
                    let p = self.0.val.array.lparray as *mut Variant;
                    let len = (self.0.val.array.rows * self.0.val.array.columns) as usize;
                    let cap = len;
                    Vec::from_raw_parts(p, len, cap);
                }
            }
            _ => {
                // nothing to do
            }
        }
    }
}

/// We need to hand-code Clone, because of the ownership issues for strings and multi.
impl Clone for Variant {
    fn clone(&self) -> Variant {
        // a simple copy is good enough for most variant types, but make sure the addin
        // is the owner
        let mut copy = Variant(self.0);
        copy.0.xltype &= !xlbitXLFree;
        copy.0.xltype |= xlbitDLLFree;

        // Special handling for string and mult, to avoid double delete of the member
        match copy.0.xltype {
            xltypeStr_xlbitDLLFree => {
                // We have a 16bit string that was originally allocated as a vector
                // but then forgotten. Reconstruct the vector, so we can clone it.
                unsafe {
                    let p = copy.0.val.str;
                    let len = *p as usize + 1;
                    let cap = len;
                    let string_vec = Vec::from_raw_parts(p, len, cap);
                    let mut cloned = string_vec.clone();
                    copy.0.val.str = cloned.as_mut_ptr();

                    // now forget everything -- we do not want either string deallocated
                    mem::forget(string_vec);
                    mem::forget(cloned);
                }
            }
            xltypeMulti_xlbitDLLFree => {
                // We have an array that was originally allocated as a vector
                // but then forgotten. Reconstruct the vector, so we can clone it.
                unsafe {
                    let p = self.0.val.array.lparray as *mut Variant;
                    let len = (self.0.val.array.rows * self.0.val.array.columns) as usize;
                    let cap = len;
                    let array = Vec::from_raw_parts(p, len, cap);
                    let mut cloned = array.clone();
                    copy.0.val.array.lparray = cloned.as_mut_ptr() as LPXLOPER12;

                    // now forget everything -- we do not want either string deallocated
                    mem::forget(array);
                    mem::forget(cloned);
                }
            }
            _ => {
                // nothing to do
            }
        }

        copy
    }
}

// NDArray support for 2d arrays
#[cfg(feature = "use_ndarray")]
use ndarray::Array2;

#[cfg(feature = "use_ndarray")]
impl<'a> From<&'a Variant> for Array2<f64> {
    fn from(v: &'a Variant) -> Array2<f64> {
        let (x, y) = v.dim();
        let mut res = Array2::zeros([y, x]);
        // Not an array
        if x == 1 && y == 1 {
            res[[0, 0]] = TryFrom::<&Variant>::try_from(v).unwrap_or_default();
        } else {
            for j in 0..y {
                for i in 0..x {
                    let index = j * x + i;
                    let slice =
                        unsafe { slice::from_raw_parts::<xloper12>(v.0.val.array.lparray, x * y) };
                    let v = slice[index];
                    res[[j, i]] = if v.xltype & xltypeMask != xltypeNum {
                        0.0f64
                    } else {
                        dbg!(&unsafe { v.val.num });
                        unsafe { v.val.num }
                    };
                }
            }
        }
        res
    }
}

#[cfg(feature = "use_ndarray")]
impl<'a> From<&'a Variant> for Array2<String> {
    fn from(v: &'a Variant) -> Array2<String> {
        let (x, y) = v.dim();
        let mut res = Array2::from_elem([y, x], String::new());
        if x == 1 && y == 1 {
            res[[0, 0]] = TryFrom::<&Variant>::try_from(v).unwrap_or_default();
        } else {
            for j in 0..y {
                for i in 0..x {
                    let index = j * x + i;
                    let slice =
                        unsafe { slice::from_raw_parts::<xloper12>(v.0.val.array.lparray, x * y) };
                    let v = slice[index];
                    res[[j, i]] = if (v.xltype & xltypeMask) != xltypeStr {
                        String::new()
                    } else {
                        let cstr_slice = unsafe {
                            let cstr: *const u16 = v.val.str;
                            let cstr_len = *cstr.offset(0) as usize;
                            slice::from_raw_parts(cstr.offset(1), cstr_len)
                        };
                        match String::from_utf16(cstr_slice) {
                            Ok(s) => s,
                            Err(e) => e.to_string(),
                        }
                    }
                }
            }
        }
        res
    }
}

#[cfg(feature = "use_ndarray")]
impl From<Array2<f64>> for Variant {
    fn from(arr: Array2<f64>) -> Variant {
        // Return as a Variant
        let mut array = arr
            .iter()
            .map(|&v| match v {
                v if v.is_nan() => Variant::from_err(xlerrNA),
                v => Variant::from(v),
            })
            .collect::<Vec<_>>();
        let rows = arr.nrows();
        let columns = arr.ncols();
        if rows == 0 || columns == 0 {
            Variant::from_err(xlerrNull)
        } else if rows > 32767 {
            Variant::from_err(xlerrValue)
        } else {
            let lparray = array.as_mut_ptr() as LPXLOPER12;
            mem::forget(array);
            Variant(XLOPER12 {
                xltype: xltypeMulti,
                val: xloper12__bindgen_ty_1 {
                    array: xloper12__bindgen_ty_1__bindgen_ty_3 {
                        lparray,
                        rows: rows as i32,
                        columns: columns as i32,
                    },
                },
            })
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[cfg(feature = "use_ndarray")]
    use ndarray::Array2;
    #[cfg(feature = "use_ndarray")]
    #[test]
    fn ndarray_conv() {
        let arr = Array2::from_shape_vec([2, 2], vec![1.0, 2.0, 3.0, 4.0]).unwrap();
        let v: Variant = From::<Array2<f64>>::from(arr.clone());
        let conv: Array2<f64> = From::<&Variant>::from(&v);
        dbg!(&arr);
        dbg!(&conv);
        assert_eq!(arr, conv);
    }
}
