//! Entry point code for xladd, based on the sample C++ code
//! supplied with the Microsoft Excel12 SDK

use crate::registrator::debug_print;
use crate::variant::Variant;
use crate::xlcall::{xlFree, xlretFailed, LPXLOPER12, XLOPER12};
use std::ffi::CStr;
use std::mem;
use std::ptr;
use widestring::U16CString;
use winapi::shared::minwindef::HMODULE;
use winapi::um::libloaderapi::{GetModuleHandleW, GetProcAddress};

const EXCEL12ENTRYPT: &[u8] = b"MdCallBack12\0";
const XLCALL32DLL: &str = "XLCall32";
const XLCALL32ENTRYPT: &[u8] = b"GetExcel12EntryPt\0";
type EXCEL12PROC = extern "stdcall" fn(
    xlfn: ::std::os::raw::c_int,
    count: ::std::os::raw::c_int,
    rgpxloper12: *const LPXLOPER12,
    xloper12res: LPXLOPER12,
) -> ::std::os::raw::c_int;
type FNGETEXCEL12ENTRYPT = extern "stdcall" fn() -> usize;

static mut XLCALL_HMODULE: HMODULE = ptr::null_mut();
static mut PEXCEL12: usize = 0;

/// Call into Excel, passing a function number as defined in xlcall and a slice
/// of Variant, and returning a Variant. To find out the number and type of
/// parameters and the expected result, please consult the Excel SDK documentation.
///
/// Note that this is a slightly inefficient call, in that it allocates a vector
/// of pointers. For example, if you have a single argument, it is faster to invoke
/// the single arg version.
pub fn excel12(xlfn: u32, opers: &mut [Variant]) -> Variant {
    debug_print(&format!("FuncID:{}, {} args)", xlfn, opers.len()));
    let mut args: Vec<LPXLOPER12> = Vec::with_capacity(opers.len());
    for oper in opers.iter_mut() {
        debug_print(&format!("arg: {}", oper));
        args.push(oper.as_mut_xloper());
    }
    let mut result = Variant::default();
    let res = excel12v(xlfn as i32, result.as_mut_xloper(), &args);
    match res {
        0 => result,
        v => {
            debug_print(&format!("ReturnCode {}", v));
            result
        }
    }
}

pub fn excel12_1(xlfn: u32, mut oper: Variant) -> Variant {
    let mut result = Variant::default();
    excel12v(xlfn as i32, result.as_mut_xloper(), &[oper.as_mut_xloper()]);
    result
}

fn fetch_excel12_entry_pt() {
    unsafe {
        if XLCALL_HMODULE.is_null() {
            let wcstr = U16CString::from_str(XLCALL32DLL).unwrap();
            XLCALL_HMODULE = GetModuleHandleW(wcstr.as_ptr());
            if !XLCALL_HMODULE.is_null() {
                let cstr = CStr::from_bytes_with_nul(XLCALL32ENTRYPT).unwrap();
                let entry_pt: usize = GetProcAddress(XLCALL_HMODULE, cstr.as_ptr()) as usize;
                if entry_pt != 0 {
                    PEXCEL12 = mem::transmute::<usize, FNGETEXCEL12ENTRYPT>(entry_pt)();
                }
            }
        }

        if PEXCEL12 == 0 {
            XLCALL_HMODULE = GetModuleHandleW(ptr::null());
            if !XLCALL_HMODULE.is_null() {
                let cstr = CStr::from_bytes_with_nul(EXCEL12ENTRYPT).unwrap();

                PEXCEL12 = GetProcAddress(XLCALL_HMODULE, cstr.as_ptr()) as usize;
            }
        }
    }
}

pub fn excel12v(xlfn: i32, oper_res: &mut XLOPER12, opers: &[LPXLOPER12]) -> i32 {
    fetch_excel12_entry_pt();

    unsafe {
        if PEXCEL12 == 0 {
            xlretFailed as i32
        } else {
            let p = opers.as_ptr();
            let len = opers.len();
            mem::transmute::<usize, EXCEL12PROC>(PEXCEL12)(xlfn, len as i32, p, oper_res)
        }
    }
}

pub fn excel_free(xloper: LPXLOPER12) -> i32 {
    fetch_excel12_entry_pt();

    unsafe {
        if PEXCEL12 == 0 {
            xlretFailed as i32
        } else {
            mem::transmute::<usize, EXCEL12PROC>(PEXCEL12)(
                xlFree as i32,
                1,
                &xloper,
                ptr::null_mut(),
            )
        }
    }
}
