// 开发单位：息县灵空网络科技工作室
// 作者：小呆呆
// 开发时间：2025-04-23
const fs = require('fs-extra');
const path = require('path');
const AdmZip = require('adm-zip');

/**
 * Excel文档处理类
 */
class ExcelProcessor {
  /**
   * 解压Excel文档
   * @param {string} xlsxPath - Excel文档路径
   * @param {string} extractPath - 解压目标路径
   * @returns {Promise<string>} - 解压后的路径
   */
  static async extract(xlsxPath, extractPath) {
    try {
      // 确保目录存在
      await fs.ensureDir(extractPath);
      
      // 创建压缩对象
      const zip = new AdmZip(xlsxPath);
      
      // 解压所有文件
      zip.extractAllTo(extractPath, true);
      
      console.log(`Excel已解压到: ${extractPath}`);
      return extractPath;
    } catch (error) {
      console.error('解压Excel文档时出错:', error);
      throw error;
    }
  }

  /**
   * 从解压的文件夹创建Excel文档
   * @param {string} folderPath - 包含Excel文档结构的文件夹路径
   * @param {string} outputPath - 输出的Excel文档路径
   * @returns {Promise<string>} - 生成的Excel文档路径
   */
  static async create(folderPath, outputPath) {
    try {
      const zip = new AdmZip();
      
      // 递归添加所有文件到压缩包
      await this._addFilesToZip(zip, folderPath, '');
      
      // 确保输出目录存在
      await fs.ensureDir(path.dirname(outputPath));
      
      // 写入zip文件
      zip.writeZip(outputPath);
      
      console.log(`Excel文档已创建: ${outputPath}`);
      return outputPath;
    } catch (error) {
      console.error('创建Excel文档时出错:', error);
      throw error;
    }
  }

  /**
   * 递归添加文件到zip
   * @private
   * @param {AdmZip} zip - AdmZip实例
   * @param {string} folderPath - 文件夹路径
   * @param {string} relativePath - 相对路径
   */
  static async _addFilesToZip(zip, folderPath, relativePath) {
    const items = await fs.readdir(folderPath);
    
    for (const item of items) {
      const itemPath = path.join(folderPath, item);
      const itemRelativePath = path.join(relativePath, item);
      
      const stats = await fs.stat(itemPath);
      
      if (stats.isDirectory()) {
        // 递归处理子目录
        await this._addFilesToZip(zip, itemPath, itemRelativePath);
      } else {
        // 添加文件到zip
        zip.addLocalFile(itemPath, path.dirname(itemRelativePath));
      }
    }
  }
}

module.exports = ExcelProcessor; 