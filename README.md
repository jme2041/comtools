# ComTools: Lightweight tools for the Component Object Model

ComTools provides the C++ includes `iptr.h`, `ubstr.h`, and `comexcept.h`,
which provide the `ComTools` namespace.

`iptr.h` is a C++ header that implements `ComTools::IPtr`, a smart pointer for
wrapping Component Object Model (COM) interfaces. `IPtr` is based on and
provides a non-Windows Runtime version of `Modern::ComPtr`. `ComPtr` was
originally developed by Kenny Kerr and released under the MIT license. `IPtr`
was originally released in its own
[repository](https://www.github.com/jme2041/iptr), but development has shifted
to this repository.

`ubstr.h` implements `ComTools::UBSTR`, which is a wrapper class for the `BSTR`
data type. `UBSTR` is based on `_UBSTR`, which was described by Don Box in
Essential COM (1998, Reading, MA: Addison-Wesley).

`comexcept.h` implements `ComTools::ComException`, which wraps `GetErrorInfo()`
to extract COM error information that was set via `SetErrorInfo()`.

ComTools is designed to be lightweight, consistent with modern C++ design
patterns, and not dependent on the ATL or MFC libraries or `__uuidof()`.

# License

Copyright (c) 2021-2022, Jeffrey M. Engelmann

`ComTools` is released under the MIT license.
For details, see [LICENSE.txt](LICENSE.txt).
