/*
rustup toolchain list
rustup toolchain add stable-i686-pc-windows-msvc
cargo +stable-i686-pc-windows-msvc build --release
*/

extern crate libc;

use libc::c_char;
use std::collections::hash_map::DefaultHasher;
use std::ffi::CStr;
use std::ffi::CString;

//use std::hash::Hash;
use std::hash::Hasher;

#[no_mangle]
pub extern "stdcall" fn FFIDivide(x: i32, y: i32) -> i32 {
    x / y
}

#[no_mangle]
pub extern "stdcall" fn FFIMultiply(x: i32, y: i32) -> i32 {
    x * y
}

#[no_mangle]
pub extern "stdcall" fn FFISubtract(x: i32, y: i32) -> i32 {
    x - y
}

#[no_mangle]
pub extern "stdcall" fn FFISum(x: i32, y: i32) -> i32 {
    x + y
}

#[no_mangle]
pub extern "stdcall" fn FFIString() -> *const c_char {
    let s = CString::new("Hello World").unwrap();
    let p = s.as_ptr();
    std::mem::forget(s);
    p
}

#[no_mangle]
//pub extern "C" fn how_many_characters(s: *const c_char) -> i32 {
pub extern "stdcall" fn how_many_characters(s: *const c_char) -> i32 {
    let c_str = unsafe {
        assert!(!s.is_null());

        CStr::from_ptr(s)
    };

    let r_str = c_str.to_str().unwrap();
    r_str.chars().count() as i32
}

#[no_mangle]
pub extern "stdcall" fn my_hasher(my_str: *const c_char) -> *const c_char {
    let c_str = unsafe {
        assert!(!my_str.is_null());

        CStr::from_ptr(my_str)
    };
    let r_str = c_str.to_str().unwrap();

    let s = DefaultHasher::new();
    let mut hasher = s.clone();
    hasher.write(r_str.as_bytes());
    let _hs = format!("{:x}", hasher.finish());

    let ss = CString::new(_hs).unwrap();
    let p = ss.as_ptr();
    std::mem::forget(ss);
    p
}