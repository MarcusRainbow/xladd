# xladd

A library to assist with the development of addins to Excel using the Excel4 and Excel12 APIs.

## The Microsoft APIs

The Excel4 API was added to Excel version 4 and is still largely supported in all subsequent versions of Excel that run on the Windows platform. (Sadly, Apple users have to make do with VBA.) In Excel 2007, an additional API was added to give support for spreadsheets with more than 256 columns and strings longer than 254 characters. The API is no longer supported in its entirety: for example, the menu modification functions no longer work. However, Microsoft have signalled strong support for the API moving forward, for example this is the only API that allows multi-threaded async operation for addins.

The API performs essentially two roles. It allows addins to control the behaviour of Excel, doing things like capturing keystrokes and forcibly writing into cells. And it allows addins to register their own functions, which can be called from Excel cells and constructed by the function wizard, just like Excel's intrinsic functions.

## The roles of xladd

xladd is a library written entirely in Rust. It is intended to greatly simplify the writing of Excel addins, with the following features:

* Wrap the functions in the Excel4 and Excel12 APIs so they can be called directly from Rust.

* Wrap the XLOPER struct that is used for communication with Excel to make it leak safe and to allow read-write access to it from safe Rust code.

* Use Rust macros to auto-derive the registration code to allow Excel to call functions written in Rust.

* Support a map-based cache to allow Rust structs to be returned from addin functions, and subsequently passed to other addin functions. XLOPER supports only basic types such as strings, numbers and arrays of strings and numbers, but a cache allows non-trivial data to be keyed by a string or number, which can thus be passed through Excel's own calculation logic.

* Allow communication between different addins, with a shared cache of structs supporting the 'any' interface.

## XLOPER and the low-level API

An XLOPER is a Microsoft-written struct containing discriminated unions, allowing it to represent the contents of an Excel cell or range of cells. It dates from Excel version 4, which predates VBA. In those days, Excel macros were a range of Excel cells, containing numbers representing calls and control flow. The control flow parts of an XLOPER are redundant now, but the struct is still used to represent the following variant types:

* Number (floating point, integer and date are all internally the same)
* String (until Excel 97, strings were limited to 254 ASCII characters, and this is one of the differences between the Excel4 and Excel12 APIs)
* Error
* Range of cells, each of which can be any of the other types
* Empty

When you register a function for Excel to invoke, you must specify the parameters of the function. They can be XLOPERs, allowing you to coerce values or reject them with your own error messages, or you can specify a few built-in types such as integers, floating point numbers or strings, in which case Excel does the coercion or rejection for you before invoking your function. The standard in the industry is to always specify XLOPERs, giving more flexibility, but in this library we give you the choice.
