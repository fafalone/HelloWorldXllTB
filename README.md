# HelloWorldXllTB
Making Excel XLL Addins in twinBASIC

![image](https://github.com/user-attachments/assets/f310f0bd-8884-44aa-9363-d10d34020b37)

This is a twinBASIC port of https://github.com/edparcell/HelloWorldXll, showing the basic concept of how to make an XLL addin using the same language VBA programmers are used to.

### Using the project
You need only to compile, then in Excel, on the Developer tab click 'Excel Add-ins', then 'Browse', and navigate to wherever your .xll file is.

### How it was made

Step 1: Create a new Standard DLL project in twinBASIC.

Step 2: Configure project settings (under the Project menu). You'll need to manually change the output path to have a `.xll` extension, since tB doesn't have a specific project type for this yet, but it's just a renamed DLL anyway. I'm not sure if it's *absolutely* needed, but since the XLCall32.dll we need won't be in the same folder as our xll, I've done like C and put it in the IAT by changing the **Project: Runtime Binding of DLL Declares** to **No**. 

Step 3: Add definitions. As part of this project, I've gone ahead and created a tB version of the entire Excel SDK's xlcall.h, so you can reuse this section in other projects and have everything you need. There's also some standard Windows API defs below this for the demo; you don't need these if you use my WinDevLib package which defines all common APIs.

Step 4: Add the `XLAutoOpen` function as a `[DllExport]` and its code:

```
    [DllExport]
    Public Function xlAutoOpen() As Integer
 
        Dim text As String = StrConv("Hello world from a twinBASIC XLL Addin!", vbFromUnicode)
        Dim text_len As Long = Len("Hello world from a twinBASIC XLL Addin!")
        Dim message As XLOPER
        message.xltype = xltypeStr
       
        Dim pStr As LongPtr = GlobalAlloc(GPTR, text_len + 2) 'Excel frees it, that's why this trouble
        CopyMemory ByVal VarPtr(message), pStr, LenB(pStr)
        CopyMemory ByVal pStr, CByte(text_len), 1
        CopyMemory ByVal pStr + 1, ByVal StrPtr(text), text_len + 1
 
        Dim dialog_type As XLOPER
        dialog_type.xltype = xltypeInt
        Dim n As Integer = 2
        CopyMemory ByVal VarPtr(dialog_type), n, 2
 
        Excel4(xlcAlert, vbNullPtr, 2, ByVal VarPtr(message), ByVal VarPtr(dialog_type))
        Return 1
    End Function
```

It's a little complicated here since we need an ANSI string that Excel can free without tB freeing it automatically as well, which would crash. So we do some memory allocation.

The biggest problem in code is the absolutely horrendous `XLOPER` type. It's got tons of unions of sub-types, neither of which tB currently supports. So I've calculated the correct size in bytes, and substituted `LongLong` because it needs to be aligned on an 8-byte boundary, even in 32bit because `Double` is a union option. No matter what the type, it's copied to offset 0, the very beginning. 

Unlike VBA, twinBASIC supports variadic functions, so we can use either `Excel4` or `Excel4v`. 

So while this project does use some twinBASIC-only syntax, using what's essentially a newer version of VBA is still far, far better than needing to know C/C++!
