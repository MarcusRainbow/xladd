#![allow(non_snake_case, non_camel_case_types, non_upper_case_globals)]

use std::{mem, fmt, slice};
use xlcall::{XLOPER12, LPXLOPER12, xloper12__bindgen_ty_1, 
    xltypeNil, xltypeInt, xltypeStr, xltypeErr, xltypeMissing, xltypeNum,
    xlbitDLLFree, xlbitXLFree,
    xlerrNull, xlerrDiv0, xlerrValue, xlerrRef, xlerrName, xlerrNum, xlerrNA, xlerrGettingData };
use entrypoint::excel_free;

const xltypeMask : u32 = !(xlbitDLLFree | xlbitXLFree);

/// Variant is a wrapper around an XLOPER12. It can contain a string, i32 or f64, or a
/// two dimensional of any mixture of these. Basically, it can contain anything that an
/// Excel cell or array of cells can contain.
pub struct Variant(XLOPER12);

impl Variant {
    /// Construct a variant containing nil. This is used in Excel to represent cells that have
    /// nothing in them. It is also a sensible starting state for an uninitialized variant.
    pub fn new() -> Variant {
        Variant(XLOPER12 { xltype : xltypeNil, val: xloper12__bindgen_ty_1 { w: 0 } })
    }

    /// Construct a variant from an LPXLOPER12, for example supplied by Excel. The assumption
    /// is that Excel continues to own the XLOPER12 and its lifetime is greater than that of
    /// the Variant we construct here. For example, the LPXLOPER may be an argument to one
    /// of our functions. We therefore do not want to own any of the data in this variant, so
    /// we clear all ownership bits. This means we treat it as a kind of dynamic mut ref. 
    pub fn from_xloper(xloper: LPXLOPER12) -> Variant {
        let mut result = Variant(unsafe { *xloper });
        result.0.xltype &= xltypeMask;    // no ownership bits
        result
    }

    /// Construct a variant containing an int (i32)
    pub fn from_int(w: i32) -> Variant {
        Variant(XLOPER12 { xltype : xltypeInt, val: xloper12__bindgen_ty_1 { w: w } })
    }

    /// Construct a variant containing a float (f64)
    pub fn from_float(num: f64) -> Variant {
        Variant(XLOPER12 { xltype : xltypeNum, val: xloper12__bindgen_ty_1 { num: num } })
    }

    /// Construct a variant containing a missing entry. This is used in function calls to
    /// signal that a parameter should be defaulted.
    pub fn missing() -> Variant {
        Variant(XLOPER12 { xltype : xltypeMissing, val: xloper12__bindgen_ty_1 { w: 0 } })
    }

    /// Construct a variant containing an error. This is used in Excel to represent standard errors
    /// that are shown as #DIV0 etc. Currently supported error codes are:
    /// xlerrNull, xlerrDiv0, xlerrValue, xlerrRef, xlerrName, xlerrNum, xlerrNA, xlerrGettingData
    pub fn from_err(xlerr: u32) -> Variant {
        Variant(XLOPER12 { xltype : xltypeErr, val: xloper12__bindgen_ty_1 { err: xlerr as i32 } })
    }

    /// Construct a variant containing a string. Strings in Excel (at least after Excel 97) are 16bit
    /// Unicode starting with a 16-bit length. The length is treated as signed, which means that
    /// strings can be no longer than 32k characters. If a string longer than this is supplied, or a 
    /// string that is not valid 16bit Unicode, an xlerrValue error is stored instead.
    pub fn from_str(s: &str) -> Variant {
        let mut wstr : Vec<u16> = s.encode_utf16().collect();
        let len = wstr.len();
        if len > 32767 {
            return Variant::from_err(xlerrValue)
        }

        // Pascal-style string with length at the start. Forget the string so we do not delete it.
        // We are now relying on the drop method of Variant to clean it up for us. Note that the
        // shrink_to_fit is essential, so the capacity is the same as the length. We have no way
        // of storing the capacity otherwise.
        wstr.insert(0, len as u16);
        wstr.shrink_to_fit();
        let p = wstr.as_mut_ptr();
        mem::forget(wstr);
  
        Variant(XLOPER12 { xltype : xltypeStr + xlbitDLLFree, val: xloper12__bindgen_ty_1 { str: p } })
    }

    /// Converts this variant to a string. Alternatively, you can use Display or to_string,
    /// which both go through this call if the variant contains a string. Guaranteed to return
    /// Some(...) if this object is of type xltypeStr. Always returns None if this object is
    /// of any other type. If the string contains a unicode string that is misformed, return
    /// the error message.
    pub fn as_string(&self) -> Option<String> {
        if (self.0.xltype & xltypeMask) != xltypeStr {
             None
        } else {
            let cstr_slice = unsafe {
                let cstr: *const u16 = self.0.val.str;
                let cstr_len = *cstr.offset(0) as usize;
                slice::from_raw_parts(cstr.offset(1), cstr_len) };
            match String::from_utf16(cstr_slice) {
                Ok(s) => Some(s),
                Err(e) => Some(e.to_string())
            }
        }
    }

    /// Converts this variant to an int. If we do not contain an int, return None. Note that
    /// Excel cells do not ever contain ints, so this would only come from a non-Excel user
    /// creating an XLOPER, for example the result of a call into Excel.
    pub fn as_i32(&self) -> Option<i32> {
        if (self.0.xltype & xltypeMask) != xltypeInt {
            None
        } else {
            Some(unsafe { self.0.val.w })
        }
    }
    
    /// Converts this variant to a float. If we do not contain a float, return None.
    pub fn as_f64(&self) -> Option<f64> {
        if (self.0.xltype & xltypeMask) != xltypeNum {
            None
        } else {
            Some(unsafe { self.0.val.num })
        }
    }

    /// Exposes the underlying XLOPER12
    pub fn as_mut_xloper(&mut self) -> &mut XLOPER12 {
        &mut self.0
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
                _ => write!(f, "#BAD_ERR")
            }
            xltypeInt => write!(f, "{}", unsafe { self.0.val.w }),
            xltypeMissing => write!(f, "#MISSING"),
            xltypeNil => write!(f, "#NIL"),
            xltypeNum => write!(f, "{}", unsafe { self.0.val.num }),
            xltypeStr => write!(f, "{}", self.as_string().unwrap()),
            _ => write!(f, "#BAD_XLOPER")
        }
    }
}

/// We need to implement Drop, as Variant is a wrapper around a union type that does
/// not know how to handle its contained pointers.
impl Drop for Variant {
    fn drop(&mut self) {
        if (self.0.xltype & xlbitXLFree) != 0 {
            excel_free(&mut self.0);
            return
        }
        match self.0.xltype {
            xltypeStr | xlbitDLLFree => {
                // We have a 16bit string that was originally allocated as a vector
                // but then forgotten. Reconstruct the vector, so its drop method
                // will clean up the memory for us.
                unsafe {
                    let p = self.0.val.str;
                    let len = *p as usize;
                    let cap = len;
                    Vec::from_raw_parts(p, len, cap);
                }
            },
            _ => {
                // nothing to do
                // TODO we need to handle arrays
            }
        }
    }
}

/// We need to hand-code Clone, because of the ownership issues for strings and multi.
impl Clone for Variant {
    fn clone(&self) -> Variant {
        // a simple copy is good enough for most variant types, but make sure the addin
        // is the owner
        let mut copy = Variant(self.0.clone());
        copy.0.xltype &= !xlbitXLFree;
        copy.0.xltype |= xlbitDLLFree;

        // Special handling for string and mult, to avoid double delete of the member
        match copy.0.xltype {
            xltypeStr | xlbitDLLFree => {

                // We have a 16bit string that was originally allocated as a vector
                // but then forgotten. Reconstruct the vector, so we can clone it.
                unsafe {
                    let p = copy.0.val.str;
                    let len = *p as usize;
                    let cap = len;
                    let string_vec = Vec::from_raw_parts(p, len, cap);
                    let mut cloned = string_vec.clone();
                    copy.0.val.str = cloned.as_mut_ptr();

                    // now forget everything -- we do not want either string deallocated
                    mem::forget(string_vec);
                    mem::forget(cloned);
                }
            },
            _ => {
                // nothing to do
                // TODO we need to handle arrays
            }
        }

        copy
    }
}