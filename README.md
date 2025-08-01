# LG Office Document Processor

一个基于Spring Boot的Office文档处理工具，支持多种文档格式的处理和翻译。

## 功能特性

- 📄 支持多种Office格式：.doc, .docx, .xls, .xlsx, .ppt, .pptx
- 🚀 三种处理引擎：Apache POI, Aspose, docx4j
- 🌐 智能翻译：中英文互译功能
- 📝 文档标记：添加翻译标记功能

## 技术栈

- Java 17
- Spring Boot 3.4.7
- Apache POI 5.2.4
- Aspose Words 23.12
- docx4j 11.4.9
- Google Translate API

## 快速开始

### 环境要求

- Java 17+
- Maven 3.6+

### 安装步骤

1. 克隆项目
```bash
git clone https://github.com/你的用户名/LG.git
cd LG
```

2. 编译项目
```bash
mvn clean compile
```

3. 运行项目
```bash
mvn spring-boot:run
```

4. 访问应用
打开浏览器访问：http://localhost:8080

## 使用说明

1. **Apache POI处理**：轻量级解决方案，支持所有格式
2. **Aspose处理**：版式保持度最高，适合复杂文档
3. **docx4j处理**：专门处理.docx格式
4. **智能翻译**：自动检测语言并进行中英文互译

## 开发指南

### 项目结构
```
src/
├── main/
│   ├── java/com/example/demo/
│   │   ├── DocumentController.java    # 主控制器
│   │   ├── TranslateService.java      # 翻译服务
│   │   └── OfficeProcessorApplication.java
│   └── resources/
│       ├── static/index.html          # 前端页面
│       └── application.properties     # 配置文件
```

### 添加新功能

1. Fork项目
2. 创建功能分支：`git checkout -b feature/新功能`
3. 提交更改：`git commit -am '添加新功能'`
4. 推送分支：`git push origin feature/新功能`
5. 创建Pull Request

## 贡献指南

欢迎提交Issue和Pull Request！

## 许可证

MIT License