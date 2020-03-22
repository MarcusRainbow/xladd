#![allow(non_snake_case, non_camel_case_types, non_upper_case_globals)]

use std::{fmt, mem, slice};
//#[cfg(feature = "try_from")]
use entrypoint::excel_free;
use std::convert::{TryFrom, TryInto};
use std::f64;
use xlcall::{
    xlbitDLLFree, xlbitXLFree, xlerrDiv0, xlerrGettingData, xlerrNA, xlerrName, xlerrNull,
    xlerrNum, xlerrRef, xlerrValue, xloper12__bindgen_ty_1, xloper12__bindgen_ty_1__bindgen_ty_3,
    xltypeBool, xltypeErr, xltypeInt, xltypeMissing, xltypeMulti, xltypeNil, xltypeNum, xltypeStr,
    LPXLOPER12, XLOPER12,
};

#[derive(Debug)]
pub enum XLAddError {
    F64ConversionFailed,
    BoolConversionFailed,
    IntConversionFailed,
    StringConversionFailed,
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

    // To float array
    pub fn convert_float_array(array: Vec<f64>, columns: usize, rows: usize) -> Variant {
        // Return as a Variant
        let mut array = array
            .iter()
            .map(|&v| match v {
                v if v.is_nan() => Variant::from_err(xlerrNA),
                v => Variant::from(v),
            })
            .collect::<Vec<_>>();
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

    // When all your values are
    pub fn convert_string_array(array: Vec<&str>, columns: usize, rows: usize) -> Variant {
        // Return as a Variant
        let mut array = array.iter().map(|v| Variant::from(*v)).collect::<Vec<_>>();
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
impl TryFrom<Variant> for i32 {
    type Error = XLAddError;
    fn try_from(v: Variant) -> Result<i32, Self::Error> {
        if (v.0.xltype & xltypeMask) != xltypeInt {
            Err(XLAddError::IntConversionFailed)
        } else {
            Ok(unsafe { v.0.val.w })
        }
    }
}

/// Converts this variant to a float. If we do not contain a float, return Err.
impl TryFrom<Variant> for f64 {
    type Error = XLAddError;
    fn try_from(v: Variant) -> Result<f64, Self::Error> {
        if (v.0.xltype & xltypeMask) != xltypeNum {
            Err(XLAddError::F64ConversionFailed)
        } else {
            Ok(unsafe { v.0.val.num })
        }
    }
}

/// Converts this variant to a bool. If we do not contain a bool, return Err.
impl TryFrom<Variant> for bool {
    type Error = XLAddError;
    fn try_from(v: Variant) -> Result<Self, Self::Error> {
        if (v.0.xltype & xltypeMask) != xltypeBool {
            Err(XLAddError::BoolConversionFailed)
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
impl TryFrom<Variant> for String {
    type Error = XLAddError;
    fn try_from(v: Variant) -> Result<Self, Self::Error> {
        if (v.0.xltype & xltypeMask) != xltypeStr {
            Err(XLAddError::StringConversionFailed)
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
impl From<Variant> for Vec<f64> {
    fn from(v: Variant) -> Vec<f64> {
        let (x, y) = v.dim();
        let mut res: Vec<f64> = Vec::new();
        for j in 0..y {
            for i in 0..x {
                res.push(v.at(i, j).try_into().map_or_else(|_| f64::NAN, |v| v));
            }
        }
        res
    }
}

/// Converts a variant into a f64 array filling the missing or invalid with f64::NAN.
/// This is so that you can handle those appropriately for your application (for example fill with the mean value or 0)
impl From<Variant> for Vec<f32> {
    fn from(v: Variant) -> Vec<f32> {
        use std::f32;
        let (x, y) = v.dim();
        let mut res: Vec<f32> = Vec::new();
        for j in 0..y {
            for i in 0..x {
                res.push(
                    v.at(i, j)
                        .try_into()
                        .map_or_else(|_| f32::NAN, |v: f64| v as f32),
                );
            }
        }
        res
    }
}

/// Converts a variant into a f64 array filling the missing or invalid with f64::NAN.
/// This is so that you can handle those appropriately for your application (for example fill with the mean value or 0)
impl From<Variant> for Vec<String> {
    fn from(v: Variant) -> Vec<String> {
        let (x, y) = v.dim();
        let mut res: Vec<String> = Vec::new();
        for j in 0..y {
            for i in 0..x {
                res.push(v.at(i, j).try_into().map_or_else(|_| String::new(), |v| v));
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

impl From<f64> for Variant {
    fn from(num: f64) -> Variant {
        Variant(XLOPER12 {
            xltype: xltypeNum,
            val: xloper12__bindgen_ty_1 { num },
        })
    }
}

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
            xltypeStr => write!(f, "{}", String::try_from(self.clone()).unwrap()),
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

/*
pub extern "stdcall" fn aarc_normalize(
    array: LPXLOPER12,
    min: LPXLOPER12,
    max: LPXLOPER12,
    scale: LPXLOPER12,
) -> LPXLOPER12 {
    match normalize(
        Variant::from(array),
        Variant::from(min),
        Variant::from(max),
        Variant::from(scale),
    ) {
        Ok(v) => LPXLOPER12::from(v),
        _ => LPXLOPER12::from(Variant::from("Invalid")),
    }
}


pub fn normalize(
    array: Variant,
    min: Variant,
    max: Variant,
    norm_type: Variant,
) -> Result<Variant, AARCError> {
    let min: f64 = min.try_into()?;
    let max: f64 = max.try_into()?;
    let norm_type: f64 = norm_type.try_into()?;
    let (x, y) = array.dim();
    let array: Vec<f64> = array.into();
    let result = match norm_type as i64 {
        1 => normalize::tanh_est(&array),
        _ => normalize::min_max_norm(&array, min, max),
    };
    Ok(Variant::convert_float_array(result, x, y))
    // Zscore normalization
    // Tanh Normalization
}

pub fn min_max_norm(array: &[f64], min: f64, max: f64) -> Vec<f64> {

*/
