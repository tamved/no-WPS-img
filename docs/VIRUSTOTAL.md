# Antivirus And VirusTotal Notes

`noWPS.exe` is a small unsigned Windows executable. Some security vendors may flag unknown unsigned tools with generic detections even when the program is not malicious.

Examples of generic heuristic labels:

```text
susgen
ml.score
malicious.moderate
trojan.generic
```

These labels do not necessarily mean the program contains a real trojan. They often mean "unknown small executable, not signed, low reputation".

## How To Verify

1. Read the source code:

```text
src/noWPS.cs
```

2. Build it locally:

```powershell
.\scripts\build.ps1
```

3. Compare behavior:

- The program opens a Windows file picker.
- It reads selected `.xlsx` or `.xlsm` files.
- It writes a new `_excel.xlsx` copy.
- It does not ask for admin rights.
- It does not use the network.
- It does not install anything.

## For Company Use

If your company blocks unsigned executables, build the project internally and sign the result with a company code-signing certificate.

The included manifest requests:

```text
asInvoker
```

so the program runs with the current user's permissions and does not request administrator rights.
