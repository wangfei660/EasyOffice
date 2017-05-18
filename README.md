## PPTMaker

This console java project is for generating **Power Point** files from a ppt template and a data file with the slides information.

PPTMaker uses a template ppt file **(default 'template.pptx')**  and a data file **(default 'data.txt')** for generating a nwe power point.

The data file must have the next structure:
```
title=my title
subtitle=my subtitle
#
title=next slide title
subtitle=next slide subtitle
#
title=and so on..
subtitle=blah blah.
# 
```

lines until hashtag are considerer as a slide.

The template must have keys for replace surrounded by @ character
```
Title: @title@
Subtitle: @subtitle@
```
### Usage
Simplest usage:
```
$ pptmaker
```

### Parameters

Getting help 
```
$ pptmaker -h
```

##### Specifying a index for template slide:

In that case the second slide will be taken as template for generating next slides 

```sh
$ pptmaker -i 1
```

##### Select custom path for either input data file or template ppt file:

The default input data file (.\data.txt) and template ppt file (.\template.pptx) path
could be pass as arguments:
```sh
$ pptmaker -d C:\custom\path\data.txt -f .\customTemplate.pptx
```
