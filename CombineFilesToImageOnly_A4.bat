@echo off
setlocal enabledelayedexpansion

:: Check if at least one argument is passed
if "%~1"=="" (
    echo Please drag and drop at least one file.
    pause
    exit /b 1
)

:: Get first file's path components
set "firstFile=%~1"
set "dp1=%~dp1"
set "n1=%~n1"

:: Set output file names based on first file
set "tempDir=%dp1%%n1%-temp"
set "finalFile=%dp1%%n1%-combined-image-only.pdf"

:: Create temp directory
if exist "!tempDir!" rd /s /q "!tempDir!"
mkdir "!tempDir!" || (
    echo Failed to create temp directory "!tempDir!". Check permissions.
    pause
    exit /b 1
)

:: Counter for processed files
set "fileCount=0"

:: Process all dragged files
:processLoop
if "%~1"=="" goto :processComplete
set /a "fileCount+=1"
set "file!fileCount!=%~1"
set "ext!fileCount!=%~x1"

:: Validate and process each file
set "isValid=0"
set "tempPdf!fileCount!=!tempDir!\file!fileCount!.pdf"
set "tempPngPrefix!fileCount!=!tempDir!\file!fileCount!-page"

:: Check file extension
for %%E in (.doc .docx .pdf .jpg .jpeg .png .bmp .gif) do (
    if /i "!ext%fileCount%!"=="%%E" set "isValid=1"
)

if !isValid! equ 0 (
    echo Skipping invalid file: "!file%fileCount%!". Supported formats: .doc, .docx, .pdf, .jpg, .jpeg, .png, .bmp, .gif
    goto :nextFile
)

:: Process based on file type
if /i "!ext%fileCount%!"==".doc" (
    echo Converting Word file "!file%fileCount%!" to PDF...
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
        "& { $word = New-Object -ComObject Word.Application; $doc = $word.Documents.Open('!file%fileCount%!',[Type]::Missing,$true); $doc.SaveAs([ref]'!tempPdf%fileCount%!',17); $doc.Close($false); $word.Quit(); }"
    if not exist "!tempPdf%fileCount%!" (
        echo Failed to convert "!file%fileCount%!" to PDF.
        goto :nextFile
    )
    call :rasterizeToPng "!tempPdf%fileCount%!" "!tempPngPrefix%fileCount%!" !fileCount!
) else if /i "!ext%fileCount%!"==".docx" (
    echo Converting Word file "!file%fileCount%!" to PDF...
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
        "& { $word = New-Object -ComObject Word.Application; $doc = $word.Documents.Open('!file%fileCount%!',[Type]::Missing,$true); $doc.SaveAs([ref]'!tempPdf%fileCount%!',17); $doc.Close($false); $word.Quit(); }"
    if not exist "!tempPdf%fileCount%!" (
        echo Failed to convert "!file%fileCount%!" to PDF.
        goto :nextFile
    )
    call :rasterizeToPng "!tempPdf%fileCount%!" "!tempPngPrefix%fileCount%!" !fileCount!
) else if /i "!ext%fileCount%!"==".pdf" (
    echo Processing PDF "!file%fileCount%!"...
    copy "!file%fileCount%!" "!tempPdf%fileCount%!" >nul
    if not exist "!tempPdf%fileCount%!" (
        echo Failed to copy "!file%fileCount%!" to temp PDF.
        goto :nextFile
    )
    call :rasterizeToPng "!tempPdf%fileCount%!" "!tempPngPrefix%fileCount%!" !fileCount!
) else (
    echo Adding image "!file%fileCount%!"...
    set "tempPngPrefix!fileCount!=!file%fileCount%!"
)

:nextFile
shift
goto :processLoop

:processComplete
if !fileCount! equ 0 (
    echo No valid files were processed.
    rd /s /q "!tempDir!"
    pause
    exit /b 1
)

:: Combine all PNGs into final image-only PDF
echo Combining all pages into image-only PDF...
set "pngList="
set "pageCount=0"
for /l %%i in (1,1,!fileCount!) do (
    if /i "!ext%%i!"==".doc" (
        for %%P in ("!tempPngPrefix%%i!*.png") do (
            set "pngList=!pngList! "%%P""
            set /a "pageCount+=1"
        )
    ) else if /i "!ext%%i!"==".docx" (
        for %%P in ("!tempPngPrefix%%i!*.png") do (
            set "pngList=!pngList! "%%P""
            set /a "pageCount+=1"
        )
    ) else if /i "!ext%%i!"==".pdf" (
        for %%P in ("!tempPngPrefix%%i!*.png") do (
            set "pngList=!pngList! "%%P""
            set /a "pageCount+=1"
        )
    ) else (
        set "pngList=!pngList! "!tempPngPrefix%%i!""
        set /a "pageCount+=1"
    )
)

if !pageCount! equ 0 (
    echo No pages were successfully processed. Temp folder preserved for debugging: "!tempDir!"
    pause
    exit /b 1
)

magick !pngList! -density 96 -quality 75 -compress jpeg -resize 794x1123 -extent 794x1123 -units PixelsPerInch "!finalFile!"

if exist "!finalFile!" (
    echo Successfully created image-only PDF: "!finalFile!" with !pageCount! pages
    for %%I in ("!finalFile!") do (
        set "fileSize=%%~zI"
        set /a "fileSizeKB=!fileSize! / 1024"
        set /a "fileSizeMB=(!fileSize! + 524288) / 1048576"
    )
    echo File size: !fileSizeMB! MB ^(!fileSizeKB! KB^)
) else (
    echo ImageMagick conversion failed.
    rd /s /q "!tempDir!"
    pause
    exit /b 1
)

:: Cleanup
rd /s /q "!tempDir!"
echo Cleanup complete.
echo Task completed.
pause
exit /b 0

:rasterizeToPng
echo Rasterizing "%~1" to PNGs...
set "outputPattern=%~2%%03d.png"
gs -sDEVICE=png16m ^
   -r300 ^
   -dNOPAUSE -dBATCH ^
   -sOutputFile="!outputPattern!" ^
   "%~1" 2> "!tempDir!\file%3-gs_error.txt"
echo Checking for PNGs in "%~2*.png"...
dir "%~2*.png" >nul 2>&1
if !errorlevel! neq 0 (
    echo Ghostscript rasterization of "%~1" failed. Check "!tempDir!\file%3-gs_error.txt" for details.
    echo No PNGs found after rasterization. Temp folder: "!tempDir!"
    dir "!tempDir!" /b
    if exist "%~2" echo Unexpected file found: "%~2"
    pause
) else (
    echo Successfully rasterized "%~1" to PNGs.
    dir "%~2*.png" /b
)
exit /b