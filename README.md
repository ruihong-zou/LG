# LG Office Document Processor V3.0

基于 Spring Boot 的多格式 Office 文档批处理与智能翻译工具。支持 Word / Excel / PPT 读取与分段抽取，结合 Kimi(Moonshot) API 进行批量稳健翻译，具备自动分批、并发、失败降级与格式保持策略。

## ✨ 核心特性
- 多格式支持：.doc .docx .xls .xlsx .ppt .pptx
- 三套解析/处理引擎：Apache POI（通用） / Aspose（版式优先） / docx4j（DOCX 专精）
- 批量翻译：自适应切批 + 并发 + 顺序稳定回写
- 智能预算控制：按 token 估算分桶，避免截断与限流
- 失败降级：并发 → 串行 → 二分拆批 → 单条兜底
- 片段跳过：空白/标点/微小 ASCII 片段直通不耗 token
- 可配置：全部阈值通过 .env 覆盖
- 模拟翻译：API 不可用时提供占位结果（调试 / 离线）

## 🧱 主要组件
```
src/main/java/com/example/demo/
├── OfficeProcessorApplication.java   # 启动
├── DocumentController.java           # 上传 / 处理入口
├── DocumentProcessor.java            # 文档解析与分段
├── TranslateService.java             # 批量翻译与分批策略核心
├── Kimi.java                         # 与 Moonshot/Kimi API 交互（健壮重试）
├── ErrorController.java              # 基础错误处理
└── model/...                         # 可选：模型/DTO
```
前端：`resources/static/index.html` 提供简易上传界面。

## 🔄 翻译流程概览
1. 解析文档 → 提取文本片段（保持原顺序）
2. 清洗 / 归一化空格
3. 预算规划：按条目数、估算 prompt+completion token、输出系数，生成批次 Range
4. 并发提交批次（受 PARALLELISM 限制），结果写入固定索引数组
5. 单批内部：过滤 trivial 片段 → JSON 打包 → Kimi API → 数量校验
6. 异常（length 截断 / 数量不符）→ 二分拆分递归重试
7. 仍失败 → 模拟翻译占位
8. 汇总写回 → 生成带标记/翻译后的新文档

## ⚙️ 环境与依赖
- Java 17
- Spring Boot 3.4.x
- Apache POI / Aspose Words / docx4j
- Lombok / Hutool / OkHttp / FastJSON
- Moonshot(Kimi) Chat Completions API

## 🚀 快速开始
```bash
git clone https://github.com/你的用户名/LG.git
cd LG

=========================================
.env.example文件使用说明
1) 复制本文件为 .env 并填写 MOONSHOT_API_KEY
2) 将 .env 加入 .gitignore（保护密钥；生产用环境注入）

本地加载示例：
  macOS/Linux:
    set -a; source .env; set +a
    ./mvn spring-boot:run

  Windows PowerShell:
Get-Content .env | ForEach-Object {
  if ($_ -match '^(.*?)=(.*)$') {
    $name=$Matches[1]; $value=$Matches[2];
    [Environment]::SetEnvironmentVariable($name,$value,'Process')
  }
}
    mvn spring-boot:run

生产环境：
  - 在部署系统（Docker/K8s/CI/云主机）里注入同名环境变量即可
  - 千万不要把真实 .env 提交到仓库
=========================================

mvn clean package
mvn spring-boot:run
# 浏览器访问：
http://localhost:8080
```

## 📁 .env 关键变量
```
MOONSHOT_API_KEY=sk-...                # 必填
MOONSHOT_API_URL=https://api.moonshot.cn/v1/chat/completions
MOONSHOT_MAX_COMPLETION=1500           # 请求中 max_tokens
MOONSHOT_RPM=5000                      # 每分钟请求上限（参考/自律）
MOONSHOT_TPM=384000                    # 每分钟 token 预算（参考）
TRANSLATE_MAX_TOKENS_PER_REQUEST=30000 # (prompt + max_tokens) 本地切批预算
TRANSLATE_MAX_ITEMS_PER_BATCH=500      # 单批最大条目
TRANSLATE_COMPLETION_MARGIN=50         # 输出安全余量
TRANSLATE_DIRECTION=ZH2EN              # 默认方向：ZH2EN / EN2ZH
TRANSLATE_TRIVIAL_PASSTHROUGH=true     # 微小片段跳过
MOONSHOT_CONCURRENCY=32                # 并发批数上限
```
调优建议：
- 大文档超时 → 降低 TRANSLATE_MAX_ITEMS_PER_BATCH
- 截断频繁 → 减少 MOONSHOT_MAX_COMPLETION 或增大 MARGIN
- QPS 受限 → 降低 MOONSHOT_CONCURRENCY

## 🔌 API（典型）
| 方法 | 路径 | 说明 |
|------|------|------|
| POST | /api/poi/process     | 使用 POI 解析处理 |
| POST | /api/aspose/process  | 使用 Aspose |
| POST | /api/docx4j/process  | 使用 docx4j |
请求通常包含：文件 + 目标语言 + 可选用户指令。

（具体字段以 `DocumentController` 实际实现为准）

## 🧪 开发
```bash
mvn -q test
```
常见脚本：
- Windows 设环境：PowerShell 读取 .env（README 顶部注释示例）
- 打包可执行 Jar：
```bash
mvn -DskipTests package
java -jar target/lg-*.jar
```

## ❓ 常见问题
1. 翻译结果顺序乱？  
  使用固定索引数组回填，若出现错位多为上游解析顺序问题。
2. 为什么有些文本未翻？  
  可能被判定为 trivial（标点/空白/单字符），可设 TRANSLATE_TRIVIAL_PASSTHROUGH=false。
3. 出现 length 截断？  
  降低单批条目或增大 TRANSLATE_COMPLETION_MARGIN，或减少 MAX_COMPLETION。
4. Aspose 相关类未加载？  
  确认已添加商业库并满足许可证策略（如需）。

## 🧩 批处理策略摘录
- 预算：估算 prompt(基础开销+输入) + 预计输出(输入估算 * 系数)
- 触发切批：条目数超限 / 预算超限 / 预计输出逼近 (MAX_COMPLETION - 安全余量)
- 出错二分：批 → 半批 → 单条 → 模拟占位

## 🤝 贡献
1. Fork
2. 创建分支：`git checkout -b feature/xxx`
3. 提交：`git commit -m "feat: xxx"`
4. 推送 & PR
5. 描述改动与测试范围

## 🪪 License
MIT

## 🔒 安全
- 不要提交真实 `.env`
- 建议生产以环境变量 / Secret 注入
- 日志避免输出敏感 Key

## 🗺️ 后续规划（示例）
- Excel/PPT 样式保留增强
- 多语言自动检测聚类
- 增量翻译缓存
- Web 前端进度与失败可视化

欢迎反馈与改进建议！
