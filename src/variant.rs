#![allow(non_snake_case, non_camel_case_types, non_upper_case_globals)]

use std::{fmt, mem, slice};
//#[cfg(feature = "try_from")]
use crate::entrypoint::excel_free;
use crate::xlcall::xloper12;
use crate::xlcall::{
    xlbitDLLFree, xlbitXLFree, xlerrDiv0, xlerrGettingData, xlerrNA, xlerrName, xlerrNull,
    xlerrNum, xlerrRef, xlerrValue, xloper12__bindgen_ty_1, xloper12__bindgen_ty_1__bindgen_ty_3,
    xltypeBool, xltypeErr, xltypeInt, xltypeMissing, xltypeMulti, xltypeNil, xltypeNum, xltypeRef,
    xltypeSRef, xltypeStr, LPXLOPER12, XLMREF12, XLOPER12, XLREF12,
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
            xltypeSRef => get_sref_dim(unsafe { &self.0.val.sref.ref_ }),
            xltypeRef => get_mref_dim(unsafe { self.0.val.mref.lpmref }),
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

    pub fn is_ref(&self) -> bool {
        let xltype = self.0.xltype & xltypeMissing;
        return xltype == xltypeRef || xltype == xltypeSRef;
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

impl From<&xloper12> for String {
    fn from(v: &xloper12) -> String {
        match v.xltype & xltypeMask {
            xltypeNum => unsafe { v.val.num }.to_string(),
            xltypeStr => {
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
            xltypeBool => unsafe { v.val.xbool == 1 }.to_string(),
            _ => String::new(),
        }
    }
}

impl From<&Variant> for String {
    fn from(v: &Variant) -> String {
        String::from(&v.0)
    }
}

impl From<&xloper12> for f64 {
    fn from(v: &xloper12) -> f64 {
        match v.xltype & xltypeMask {
            xltypeNum => unsafe { v.val.num },
            xltypeInt => unsafe { v.val.w as f64 },
            xltypeStr => 0.0,
            xltypeBool => (unsafe { v.val.xbool == 1 }) as i64 as f64,
            _ => 0.0,
        }
    }
}

impl From<&Variant> for f64 {
    fn from(v: &Variant) -> f64 {
        f64::from(&v.0)
    }
}

impl From<&xloper12> for bool {
    fn from(v: &xloper12) -> bool {
        match v.xltype & xltypeMask {
            xltypeNum => unsafe { v.val.num != 0.0 },
            xltypeStr => false,
            xltypeBool => unsafe { v.val.xbool != 0 },
            _ => false,
        }
    }
}

impl From<&Variant> for bool {
    fn from(v: &Variant) -> bool {
        bool::from(&v.0)
    }
}

/// Converts a variant into a f64 array filling the missing or invalid with f64::NAN.
/// This is so that you can handle those appropriately for your application (for example fill with the mean value or 0)
impl<'a> From<&'a Variant> for Vec<f64> {
    fn from(v: &'a Variant) -> Vec<f64> {
        let (x, y) = v.dim();
        let mut res = Vec::with_capacity(x * y);
        if x == 1 && y == 1 {
            res.push(f64::from(v))
        } else {
            res.resize(x * y, 0.0);
            let slice = unsafe { slice::from_raw_parts::<xloper12>(v.0.val.array.lparray, x * y) };
            for j in 0..y {
                for i in 0..x {
                    let index = j * x + i;
                    let v = slice[index];
                    res[index] = f64::from(&v);
                }
            }
        }
        res
    }
}

/// Converts a variant into a string array filling the missing or invalid with f64::NAN.
/// This is so that you can handle those appropriately for your application (for example fill with the mean value or 0)
impl<'a> From<&'a Variant> for Vec<String> {
    fn from(v: &'a Variant) -> Vec<String> {
        let (x, y) = v.dim();
        let mut res = Vec::with_capacity(x * y);
        if x == 1 && y == 1 {
            res.push(String::from(v));
        } else {
            res.resize(x * y, String::new());
            let slice = unsafe { slice::from_raw_parts::<xloper12>(v.0.val.array.lparray, x * y) };
            for j in 0..y {
                for i in 0..x {
                    let index = j * x + i;
                    let v = slice[index];
                    res[index] = String::from(&v);
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

// For async functions
#[derive(Debug)]
pub struct XLOPERPtr(pub *mut xloper12);
unsafe impl Send for XLOPERPtr {}

impl From<XLOPERPtr> for Variant {
    fn from(xloper: XLOPERPtr) -> Variant {
        Variant(unsafe { std::mem::transmute::<XLOPER12, xloper12>(*xloper.0) })
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
        match num {
            num if num.is_nan() => Variant::from_err(xlerrNA),
            num if num.is_infinite() => Variant::from_err(xlerrNA),
            num => Variant(XLOPER12 {
                xltype: xltypeNum,
                val: xloper12__bindgen_ty_1 { num },
            }),
        }
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

/// Construct a variant containing a string. Strings in Excel (at least after Excel 97) are 16bit
/// Unicode starting with a 16-bit length. The length is treated as signed, which means that
/// strings can be no longer than 32k characters. If a string longer than this is supplied, or a
/// string that is not valid 16bit Unicode, an xlerrValue error is stored instead.
impl From<&str> for Variant {
    fn from(s: &str) -> Variant {
        let mut wstr: Vec<u16> = s.encode_utf16().collect();
        if wstr.len() > 65534 {
            return Variant::from_err(xlerrValue);
        }
        // Pascal-style string with length at the start. Forget the string so we do not delete it.
        // We are now relying on the drop method of Variant to clean it up for us. Note that the
        // shrink_to_fit is essential, so the capacity is the same as the length. We have no way
        // of storing the capacity otherwise.
        wstr.insert(0, wstr.len() as u16);
        wstr.shrink_to_fit();
        let p = wstr.as_mut_ptr();
        mem::forget(wstr);
        Variant(XLOPER12 {
            xltype: xltypeStr | xlbitDLLFree,
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
        let mut wstr: Vec<u16> = s.encode_utf16().collect();
        if wstr.len() > 65534 {
            return Variant::from_err(xlerrValue);
        }
        // Pascal-style string with length at the start. Forget the string so we do not delete it.
        // We are now relying on the drop method of Variant to clean it up for us. Note that the
        // shrink_to_fit is essential, so the capacity is the same as the length. We have no way
        // of storing the capacity otherwise.
        wstr.insert(0, wstr.len() as u16);
        wstr.shrink_to_fit();
        let p = wstr.as_mut_ptr();
        mem::forget(wstr);
        Variant(XLOPER12 {
            xltype: xltypeStr | xlbitDLLFree,
            val: xloper12__bindgen_ty_1 { str: p },
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
        } else if rows == 1 && columns == 1 {
            Variant::from(arr.0[0])
        } else {
            Variant(XLOPER12 {
                xltype: xltypeMulti | xlbitDLLFree,
                val: xloper12__bindgen_ty_1 {
                    array: xloper12__bindgen_ty_1__bindgen_ty_3 {
                        lparray,
                        rows: std::cmp::min(1_048_575, rows as i32),
                        columns: std::cmp::min(16383, columns as i32),
                    },
                },
            })
        }
    }
}

// Construct 2d variant array from (string,f64)
impl From<Vec<(String, f64)>> for Variant {
    fn from(arr: Vec<(String, f64)>) -> Variant {
        let mut array = Vec::new();
        arr.iter().for_each(|v| {
            array.push(Variant::from(v.0.as_str()));
            array.push(Variant::from(v.1))
        });

        let lparray = array.as_mut_ptr() as LPXLOPER12;
        mem::forget(array);
        let rows = arr.len();
        let columns = 2;
        if rows == 0 || columns == 0 {
            Variant::from_err(xlerrNull)
        } else {
            Variant(XLOPER12 {
                xltype: xltypeMulti | xlbitDLLFree,
                val: xloper12__bindgen_ty_1 {
                    array: xloper12__bindgen_ty_1__bindgen_ty_3 {
                        lparray,
                        rows: std::cmp::min(1_048_575, rows as i32),
                        columns: std::cmp::min(16383, columns as i32),
                    },
                },
            })
        }
    }
}

impl From<Vec<&str>> for Variant {
    fn from(arr: Vec<&str>) -> Variant {
        let mut array = Vec::new();
        arr.iter().for_each(|&v| {
            array.push(Variant::from(v));
        });

        let lparray = array.as_mut_ptr() as LPXLOPER12;
        mem::forget(array);
        let rows = 1;
        let columns = arr.len();
        if rows == 0 || columns == 0 {
            Variant::from_err(xlerrNull)
        } else {
            Variant(XLOPER12 {
                xltype: xltypeMulti | xlbitDLLFree,
                val: xloper12__bindgen_ty_1 {
                    array: xloper12__bindgen_ty_1__bindgen_ty_3 {
                        lparray,
                        rows: std::cmp::min(1_048_575, rows as i32),
                        columns: std::cmp::min(16383, columns as i32),
                    },
                },
            })
        }
    }
}

/// Pass in a tuple of (array, columns), it will calculate the number of rows.
impl From<&(&[f64], usize)> for Variant {
    fn from(arr: &(&[f64], usize)) -> Variant {
        // Return as a Variant
        let mut array = arr.0.iter().map(|&v| Variant::from(v)).collect::<Vec<_>>();
        let rows = if arr.1 == 0 { 0 } else { arr.0.len() / arr.1 };
        let columns = arr.1;
        if rows == 0 || columns == 0 {
            Variant::from_err(xlerrNull)
        } else if rows == 1 && columns == 1 {
            Variant::from(arr.0[0])
        } else {
            let lparray = array.as_mut_ptr() as LPXLOPER12;
            mem::forget(array);
            Variant(XLOPER12 {
                xltype: xltypeMulti | xlbitDLLFree,
                val: xloper12__bindgen_ty_1 {
                    array: xloper12__bindgen_ty_1__bindgen_ty_3 {
                        lparray,
                        rows: std::cmp::min(1_048_575, rows as i32),
                        columns: std::cmp::min(16383, columns as i32),
                    },
                },
            })
        }
    }
}

// Gets the array size of a multi-cell reference. If the reference is badly formed,
// returns (0, 0)
fn get_mref_dim(mref: *const XLMREF12) -> (usize, usize) {
    // currently we only handle single contiguous references
    if mref.is_null() || unsafe { (*mref).count } != 1 {
        return (0, 0);
    }

    return get_sref_dim(unsafe { &(*mref).reftbl[0] });
}

// Gets the array size of a single-cell reference
fn get_sref_dim(sref: &XLREF12) -> (usize, usize) {
    let rows = 1 + (sref.rwLast - sref.rwFirst) as usize;
    let cols = 1 + (sref.colLast - sref.colFirst) as usize;
    (cols, rows)
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
                v => write!(f, "#BAD_ERR {}", v),
            },
            xltypeInt => write!(f, "{}", unsafe { self.0.val.w }),
            xltypeMissing => write!(f, "#MISSING"),
            xltypeMulti => write!(f, "#MULTI"),
            xltypeNil => write!(f, "#NIL"),
            xltypeNum => write!(f, "{}", unsafe { self.0.val.num }),
            xltypeStr => write!(f, "{}", String::try_from(&self.clone()).unwrap()),
            xlerrNull => write!(f, "#NULL"),
            v => write!(f, "#BAD_ERR {}", v),
        }
    }
}

impl fmt::Debug for Variant {
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
                v => write!(f, "#BAD_ERR {}", v),
            },
            xlerrNull => write!(f, "#NULL"),
            xltypeInt => write!(f, "{}", unsafe { self.0.val.w }),
            xltypeMissing => write!(f, "#MISSING"),
            xltypeMulti => write!(f, "#MULTI"),
            xltypeNil => write!(f, "#NIL"),
            xltypeNum => write!(f, "{}", unsafe { self.0.val.num }),
            xltypeStr => write!(f, "{}", String::try_from(&self.clone()).unwrap()),
            v => write!(f, "#BAD_XLOPER {}", v),
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
            res[[0, 0]] = f64::from(v);
        } else {
            let slice = unsafe { slice::from_raw_parts::<xloper12>(v.0.val.array.lparray, x * y) };
            for j in 0..y {
                for i in 0..x {
                    let index = j * x + i;
                    let v = slice[index];
                    res[[j, i]] = f64::from(&v);
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
            res[[0, 0]] = String::from(v);
        } else {
            let slice = unsafe { slice::from_raw_parts::<xloper12>(v.0.val.array.lparray, x * y) };
            for j in 0..y {
                for i in 0..x {
                    let index = j * x + i;
                    let v = slice[index];
                    res[[j, i]] = String::from(&v);
                }
            }
        }
        res
    }
}
#[cfg(feature = "use_ndarray")]
impl From<Array2<Variant>> for Variant {
    fn from(arr: Array2<Variant>) -> Variant {
        let mut array = arr.iter().map(|v| v.clone()).collect::<Vec<_>>();
        let lparray = array.as_mut_ptr() as LPXLOPER12;
        let rows = arr.nrows();
        let columns = arr.ncols();
        if rows == 0 || columns == 0 {
            Variant::from_err(xlerrNull)
        } else {
            mem::forget(array);
            Variant(XLOPER12 {
                xltype: xltypeMulti | xlbitDLLFree,
                val: xloper12__bindgen_ty_1 {
                    array: xloper12__bindgen_ty_1__bindgen_ty_3 {
                        lparray,
                        rows: std::cmp::min(1_048_575, rows as i32),
                        columns: std::cmp::min(16383, columns as i32),
                    },
                },
            })
        }
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
                v if v.is_infinite() => Variant::from_err(xlerrNA),
                v => Variant::from(v),
            })
            .collect::<Vec<_>>();
        let rows = arr.nrows();
        let columns = arr.ncols();
        if rows == 0 || columns == 0 {
            Variant::from_err(xlerrNull)
        } else if rows == 1 && columns == 1 {
            Variant::from(arr[[0, 0]])
        } else {
            let lparray = array.as_mut_ptr() as LPXLOPER12;
            mem::forget(array);
            Variant(XLOPER12 {
                xltype: xltypeMulti | xlbitDLLFree,
                val: xloper12__bindgen_ty_1 {
                    array: xloper12__bindgen_ty_1__bindgen_ty_3 {
                        lparray,
                        rows: std::cmp::min(1_048_575, rows as i32),
                        columns: std::cmp::min(16383, columns as i32),
                    },
                },
            })
        }
    }
}

#[cfg(feature = "use_ndarray")]
impl From<Array2<String>> for Variant {
    fn from(arr: Array2<String>) -> Variant {
        // Return as a Variant
        let mut array = arr
            .iter()
            .map(|v| match v {
                v if v.is_empty() => Variant::from_err(xlerrNA),
                v => Variant::from(v.as_str()),
            })
            .collect::<Vec<_>>();
        let rows = arr.nrows();
        let columns = arr.ncols();
        if rows == 0 || columns == 0 {
            Variant::from_err(xlerrNull)
        } else if rows == 1 && columns == 1 {
            Variant::from(arr[[0, 0]].as_str())
        } else {
            let lparray = array.as_mut_ptr() as LPXLOPER12;
            mem::forget(array);
            Variant(XLOPER12 {
                xltype: xltypeMulti | xlbitDLLFree,
                val: xloper12__bindgen_ty_1 {
                    array: xloper12__bindgen_ty_1__bindgen_ty_3 {
                        lparray,
                        rows: std::cmp::min(1_048_575, rows as i32),
                        columns: std::cmp::min(16383, columns as i32),
                    },
                },
            })
        }
    }
}

#[macro_export]
macro_rules! max_col {
    ($x: expr) => ($x);
    ($x: expr, $($z: expr),+) => (::std::cmp::max($x, max_col!($($z),*)));
}

#[cfg(feature = "use_ndarray")]
#[macro_export]
macro_rules! make_row_table {
    ($row:expr,$name:literal,$val:expr $(,$rest_name:literal,$rest_val:expr)*) => {{
        let max_col = max_col!($val.len() $(,$rest_val.len())*);
        let mut res = Array2::default([$row,max_col+1]);
        res.fill(Variant::from(""));
        res[[0,0]] = Variant::from($name);
        $val.iter().enumerate().for_each(|(idx,&v)| res[[0,idx+1]] = Variant::from(v));
        let row_id = 0;
        $(
            let row_id = row_id + 1;
            res[[row_id,0]] = Variant::from($rest_name);
            $rest_val.iter().enumerate().for_each(|(idx,&v)| res[[row_id,idx+1]] = Variant::from(v));
        )*
        Variant::from(res)
    }};
}

#[cfg(feature = "use_ndarray")]
#[macro_export]
macro_rules! make_col_table {
    ($col:expr,$name:literal,$val:expr $(,$rest_name:literal,$rest_val:expr)*) => {{
        let max_row = max_col!($val.len() $(,$rest_val.len())*);
        let mut res = Array2::default([max_row+1,$col]);
        res.fill(Variant::from(""));
        res[[0,0]] = Variant::from($name);
        $val.iter().enumerate().for_each(|(idx,&v)| res[[idx+1,0]] = Variant::from(v));
        let col_id = 0;
        $(
            let col_id = col_id + 1;
            res[[0,col_id]] = Variant::from($rest_name);
            $rest_val.iter().enumerate().for_each(|(idx,&v)| res[[idx+1,col_id]] = Variant::from(v));
        )*
        Variant::from(res)
    }};
}

#[macro_export]
macro_rules! check_arr {
    ($name:ident) => {
        if $name.is_empty() {
            error!("{} At least 1 value reqd", stringify!($arg));
            return Err(format!("{} At least 1 value reqd", stringify!($arg)).into());
        }
    };
}

#[macro_export]
macro_rules! check_arr_non_zero_nan {
    ($name:ident) => {
        if $name.is_empty() {
            error!("{} At least 1 value reqd", stringify!($arg));
            return Err(format!("{} At least 1 value reqd", stringify!($arg)).into());
        }
        if $name.iter().any(|&v| v.is_nan() || v == 0.0) {
            error!("{} Range has a unexpected nan or 0.0", stringify!($arg));
            return Err(format!("{} Range has a unexpected nan or 0.0", stringify!($arg)).into());
        }
    };
}

#[macro_export]
macro_rules! check_float {
    ($name:ident $cmp:tt $v:tt $(,$rname:ident $rcmp:tt $rv:tt)*) => {
        if !($name $cmp $v) {
            error!(
                "{} Floating point value must be {} {}",
                stringify!($name),
                stringify!($cmp),
                stringify!($v)
            );

            return Err(format!(
                "{} Floating point value must be {} {}",
                stringify!($name),
                stringify!($cmp),
                stringify!($v)
            )
            .into());
        }
        $(
            if !($rname $rcmp $rv) {
                error!(
                    "{} Floating point value must be {} {}",
                    stringify!($rname),
                    stringify!($rcmp),
                    stringify!($rv)
                );
                return Err(format!(
                "{} Floating point value must be {} {}",
                stringify!($rname),
                stringify!($rcmp),
                stringify!($rv)
                )
                .into());
            }
        )*
    };
}

#[macro_export]
macro_rules! check_str {
    ($name:ident $(,$rest:ident)*) => {
        if $name.is_empty() {
            error!("{} At least 1 value reqd",stringify!($name));
            return Err(format!("{} At least 1 value reqd", stringify!($name)).into());
        }
        $(
            if $rest.is_empty() {
                error!("{} At least 1 value reqd",stringify!($rest));
                return Err(format!("{} At least 1 value reqd", stringify!($rest)).into());
            }
        )*
    };
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
