// 开发单位：息县灵空网络科技工作室
// 作者：小呆呆
// 开发时间：2025-04-23

const WordProcessor = require('./lib/wordProcessor');
const ExcelProcessor = require('./lib/excelProcessor');
const PowerPointProcessor = require('./lib/pptProcessor');
const path = require('path');

/**
 * Office文档处理库
 */
class OfficeDocProcessor {
  /**
   * 解压Word文档
   * @param {string} docxPath - Word文档路径
   * @param {string} extractPath - 解压目标路径
   * @returns {Promise<string>} - 解压后的路径
   */
  static async extractWord(docxPath, extractPath) {
    return WordProcessor.extract(docxPath, extractPath);
  }

  /**
   * 从解压的文件夹创建Word文档
   * @param {string} folderPath - 包含Word文档结构的文件夹路径
   * @param {string} outputPath - 输出的Word文档路径
   * @returns {Promise<string>} - 生成的Word文档路径
   */
  static async createWord(folderPath, outputPath) {
    return WordProcessor.create(folderPath, outputPath);
  }

  /**
   * 解压Excel文档
   * @param {string} xlsxPath - Excel文档路径
   * @param {string} extractPath - 解压目标路径
   * @returns {Promise<string>} - 解压后的路径
   */
  static async extractExcel(xlsxPath, extractPath) {
    return ExcelProcessor.extract(xlsxPath, extractPath);
  }

  /**
   * 从解压的文件夹创建Excel文档
   * @param {string} folderPath - 包含Excel文档结构的文件夹路径
   * @param {string} outputPath - 输出的Excel文档路径
   * @returns {Promise<string>} - 生成的Excel文档路径
   */
  static async createExcel(folderPath, outputPath) {
    return ExcelProcessor.create(folderPath, outputPath);
  }

  /**
   * 解压PowerPoint文档
   * @param {string} pptxPath - PowerPoint文档路径
   * @param {string} extractPath - 解压目标路径
   * @returns {Promise<string>} - 解压后的路径
   */
  static async extractPowerPoint(pptxPath, extractPath) {
    return PowerPointProcessor.extract(pptxPath, extractPath);
  }

  /**
   * 从解压的文件夹创建PowerPoint文档
   * @param {string} folderPath - 包含PowerPoint文档结构的文件夹路径
   * @param {string} outputPath - 输出的PowerPoint文档路径
   * @returns {Promise<string>} - 生成的PowerPoint文档路径
   */
  static async createPowerPoint(folderPath, outputPath) {
    return PowerPointProcessor.create(folderPath, outputPath);
  }
}

module.exports = OfficeDocProcessor; 