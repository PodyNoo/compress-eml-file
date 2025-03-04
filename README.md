# compress-eml-file

Tool to optimize .eml files size by converting found images (from attachements) to jpeg and reducing their quality (works also on attached .zip and .pdf)

to use : `.\compress-eml-file [path containing eml files]`, this will create `output/` folder

Warning : Be careful with result of .pdf and .zip contents

Using :
* [MimeKit](https://github.com/jstedfast/MimeKit) to manipulate eml files
* [QPdfNet](https://github.com/Sicos1977/QPdfNet) to optimize pdf sizes
* [SharpZipLib](https://github.com/icsharpcode/SharpZipLib) to manipulate zip files
* [SkiaSharp](https://github.com/mono/SkiaSharp) to convert / reduce quality of images (except for .pdf, its managed by QPdfNet)