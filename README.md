# Office文档处理库

一个基于JavaScript的库，用于处理Word、Excel和PowerPoint文档，支持在Node.js环境下运行。

## 功能特点

- 通过更改文件后缀为zip并解压，处理Word、Excel和PowerPoint文档
- 通过将文件压缩为docx、xlsx和pptx后缀文件，生成Word、Excel和PowerPoint文档
- 简单易用的API接口
- 支持异步操作

## 安装

```bash
npm install lsky-office-transition
```

## 使用方法

### 解压Word文档

```javascript
const OfficeDocProcessor = require('lsky-office-transition');

// 解压Word文档
const docxPath = 'path/to/document.docx';
const extractPath = 'path/to/extracted/folder';

OfficeDocProcessor.extractWord(docxPath, extractPath)
  .then(path => {
    console.log(`文档已解压到: ${path}`);
  })
  .catch(error => {
    console.error('解压出错:', error);
  });
```

### 创建Word文档

```javascript
const OfficeDocProcessor = require('lsky-office-transition');

// 创建Word文档
const folderPath = 'path/to/document/structure';
const outputPath = 'path/to/output/document.docx';

OfficeDocProcessor.createWord(folderPath, outputPath)
  .then(path => {
    console.log(`文档已创建: ${path}`);
  })
  .catch(error => {
    console.error('创建出错:', error);
  });
```

### 解压Excel文档

```javascript
const OfficeDocProcessor = require('lsky-office-transition');

// 解压Excel文档
const xlsxPath = 'path/to/spreadsheet.xlsx';
const extractPath = 'path/to/extracted/folder';

OfficeDocProcessor.extractExcel(xlsxPath, extractPath)
  .then(path => {
    console.log(`Excel已解压到: ${path}`);
  })
  .catch(error => {
    console.error('解压出错:', error);
  });
```

### 创建Excel文档

```javascript
const OfficeDocProcessor = require('lsky-office-transition');

// 创建Excel文档
const folderPath = 'path/to/spreadsheet/structure';
const outputPath = 'path/to/output/spreadsheet.xlsx';

OfficeDocProcessor.createExcel(folderPath, outputPath)
  .then(path => {
    console.log(`Excel已创建: ${path}`);
  })
  .catch(error => {
    console.error('创建出错:', error);
  });
```

### 解压PowerPoint文档

```javascript
const OfficeDocProcessor = require('lsky-office-transition');

// 解压PowerPoint文档
const pptxPath = 'path/to/presentation.pptx';
const extractPath = 'path/to/extracted/folder';

OfficeDocProcessor.extractPowerPoint(pptxPath, extractPath)
  .then(path => {
    console.log(`PowerPoint已解压到: ${path}`);
  })
  .catch(error => {
    console.error('解压出错:', error);
  });
```

### 创建PowerPoint文档

```javascript
const OfficeDocProcessor = require('lsky-office-transition');

// 创建PowerPoint文档
const folderPath = 'path/to/presentation/structure';
const outputPath = 'path/to/output/presentation.pptx';

OfficeDocProcessor.createPowerPoint(folderPath, outputPath)
  .then(path => {
    console.log(`PowerPoint已创建: ${path}`);
  })
  .catch(error => {
    console.error('创建出错:', error);
  });
```

## API参考

### WordProcessor

#### `extractWord(docxPath, extractPath)`

- `docxPath` (string): Word文档路径
- `extractPath` (string): 解压目标路径
- 返回: Promise<string> - 解压后的路径

#### `createWord(folderPath, outputPath)`

- `folderPath` (string): 包含Word文档结构的文件夹路径
- `outputPath` (string): 输出的Word文档路径
- 返回: Promise<string> - 生成的Word文档路径

### ExcelProcessor

#### `extractExcel(xlsxPath, extractPath)`

- `xlsxPath` (string): Excel文档路径
- `extractPath` (string): 解压目标路径
- 返回: Promise<string> - 解压后的路径

#### `createExcel(folderPath, outputPath)`

- `folderPath` (string): 包含Excel文档结构的文件夹路径
- `outputPath` (string): 输出的Excel文档路径
- 返回: Promise<string> - 生成的Excel文档路径

### PowerPointProcessor

#### `extractPowerPoint(pptxPath, extractPath)`

- `pptxPath` (string): PowerPoint文档路径
- `extractPath` (string): 解压目标路径
- 返回: Promise<string> - 解压后的路径

#### `createPowerPoint(folderPath, outputPath)`

- `folderPath` (string): 包含PowerPoint文档结构的文件夹路径
- `outputPath` (string): 输出的PowerPoint文档路径
- 返回: Promise<string> - 生成的PowerPoint文档路径

## 注意事项

- 需要确保有对相应文件和目录的读写权限
- 确保输出路径的目录已存在或有权限创建
- Node.js版本要求: >= 12.0.0

## 许可证
[Apache License 2.0](https://github.com/xsxdd/lsky-office-transition?tab=Apache-2.0-1-ov-file)
