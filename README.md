﻿## JpegToWord

This is a console application merging image files into one Word document.

### Usage:

```
cd .\JpegToWord\bin\Debug\netcoreapp3.1\
```

```
JpegToWord [options]
```

### Options:
```
  --images <images>            Specify paths to your incoming images
  --imageFolder <imageFolder>  Specify path to your folder containing incoming images
  --filename <filename>        Name for output Word file [default: MergedFile{DateTime.Now:yyMMddHHmmssff}]
  --output <output>            Path to directory where the output Word will be created [default: C:\Users\...\...\Desktop]
  --run <run>                  Specify true if want to run file after creation, default is false
  --spacing <spacing>          Specify spacing between images. Allowed value 0 - 100px, default is null
  --header <header>            Specify path to your Json file, default is null
  --version                    Show version information
  -?, -h, --help               Show help and usage information
  ```


