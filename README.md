# New Xlsx-Loader
https://github.com/NanameHacha/html-excel-loader

# Xlsx-Loader
Load xlsx files directly in HTML.

### How to use

Use a single table on the one page

```
<script type="text/javascript" src="/core.js"></script>
<script type="text/javascript" src="/main.js"></script>
<xlsx-render content="/your/path/to/excel/file.xlsx"></xlsx-render>
<div id="result"></div>
```

Use multiple tables on the one page

```
<script type="text/javascript" src="/core.js"></script>
<script type="text/javascript" src="/num/01.js"></script>
<script type="text/javascript" src="/num/02.js"></script>
<xlsx-render01 content="/your/path/to/excel/file01.xlsx"></xlsx-render01>
<xlsx-render02 content="/your/path/to/excel/file01.xlsx"></xlsx-render02>
<div id="result01"></div>
<div id="result02"></div>
```
