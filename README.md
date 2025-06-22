# Report Generator Backend

## 项目描述

这是一个基于 Spring Boot 和 Apache POI 构建的后端服务，专门为前端的日志分析仪表盘项目提供报告生成功能。它被设计为一个“哑处理器”，不包含复杂的业务逻辑，完全依赖于前端提供的 Excel 模板文件、解析后的日志数据以及详细的映射规则来动态生成 Excel 报告。

本项目的主要目标是接管报告生成中与文件操作相关的复杂任务，例如加载 Excel 模板、根据指令填充数据、复制 Sheet 结构（包括样式和图片等）以及处理不同的导出格式（单表、多 Sheet、ZIP 包），从而减轻前端的负担并利用 Java 在文件处理方面的优势。

## 主要功能

*   接收前端通过 HTTP POST 请求上传的 Excel 模板文件。
*   接收前端发送的 JSON 数据，包含待填充的日志记录列表和映射规则。
*   支持多种导出模式：
    *   **Single Sheet:** 将所有选中的日志记录的数据，根据映射规则，按列偏移填充到模板的第一个 Sheet 中。
    *   **Multi-Sheet:** 为每一条选中的日志记录，创建一个新的 Sheet（基于模板第一个 Sheet 的副本），并填充该记录的数据。新 Sheet 的名称通常基于记录的 SN。所有 Sheet 合并在一个 Excel 文件中。
    *   **ZIP Files:** 为每一条选中的日志记录，生成一个独立的 Excel 文件（基于模板第一个 Sheet 的副本并填充数据），然后将所有生成的 Excel 文件压缩成一个 ZIP 包。
*   复制模板 Sheet 的内容，包括：
    *   单元格值和类型。
    *   大部分单元格样式（通过样式映射机制，以提高颜色和字体的准确性）。
    *   合并单元格区域。
    *   列宽。
    *   基础图片（限于 XSSFClientAnchor 类型的图片）。
*   将生成的 Excel 文件或 ZIP 包作为 HTTP 响应流返回给前端。

## 技术栈

*   **后端框架:** Spring Boot 3.x
*   **Java 版本:** Java 17+
*   **Excel 处理:** Apache POI 5.x
*   **构建工具:** Maven
*   **依赖管理:** Spring Boot Starter Parent
*   **开发便利:** Lombok (可选，用于简化 JavaBean 代码)

## 开发环境要求

*   JDK 17 或更高版本
*   Maven 3.x
*   IntelliJ IDEA, Eclipse 或其他兼容 Spring Boot 的 IDE (推荐 IntelliJ IDEA)

## 快速开始

### 1. 克隆项目

如果您已经将项目代码托管到版本控制系统，请克隆项目：

```bash
git clone <您的仓库地址>
cd report-generator-backend
Use code with caution.
Markdown
如果您是直接创建的，请确保项目结构和文件（特别是 pom.xml）已准备好。
2. 导入到 IDE
使用您偏好的 IDE（如 IntelliJ IDEA）导入项目。对于 Maven 项目，IDE 通常会自动检测 pom.xml 并下载所有依赖项。请确保 Maven 同步成功，没有依赖下载失败的错误。
3. 构建项目
打开项目根目录的终端或使用 IDE 集成终端，运行 Maven 命令进行清理和构建：
Generated bash
mvn clean install
Use code with caution.
Bash
确保构建成功 (BUILD SUCCESS)。
运行应用
本项目是一个 Spring Boot 应用，可以通过多种方式运行：
方法一：通过 IntelliJ IDEA 运行 (推荐)
打开 src/main/java/com/obsidian/reportgeneratorbackend/ReportGeneratorBackendApplication.java 文件。
找到 main 方法。
点击方法左侧的绿色运行按钮，选择 Run 'ReportGeneratorBackendApplication'。
在 IDE 底部的 Run 窗口查看日志输出，直到看到类似 Tomcat started on port(s): 8080 的信息，表示后端已成功启动。
方法二：通过 Maven 命令行运行
打开项目根目录的终端或使用 IDE 集成终端。
运行命令：
Generated bash
mvn spring-boot:run
Use code with caution.
Bash
等待应用启动，观察终端输出的日志。
默认情况下，后端会在 8080 端口 启动。如果需要修改端口，请在 src/main/resources/application.properties 文件中添加 server.port=<新的端口号>。
API 文档
本项目提供一个 RESTful API 端点用于报告生成。
1. 报告生成端点
URL: /api/reports/generate
方法: POST
Content-Type: multipart/form-data
2. 请求体 (multipart/form-data)
请求体必须包含两个部分：
Part 1: Excel 模板文件
Name: template
Type: File (例如 MultipartFile 在 Spring MVC 中接收)
Description: 前端用户上传的原始 Excel 模板文件字节。
Part 2: 包含生成指令和数据的 JSON 对象
Name: request
Type: application/json (作为 Blob 发送的 JSON 字符串)
Description: 封装了导出模式、映射规则和日志数据。对应的 Java DTO 是 com.obsidian.reportgeneratorbackend.dto.ReportGenerationRequest。
request 部分的 JSON 结构 (ReportGenerationRequest):
Generated json
{
  "exportMode": "single-sheet" | "multi-sheet" | "zip-files", // 导出模式，必须是其中之一
  "mappingRules": { // 映射规则对象
    "[SN] (序列号)": { // 源数据项名称，例如 "[SN] (序列号)" 或某个具体的测试项名称
      "address": "4_2", // 目标单元格地址，格式 "行索引_列索引" (0-based)
      "unit": "V",      // 可选，附加的单位
      "decimals": 2     // 可选，保留的小数位数 (例如 2)
    },
    "电池电压": {
      "address": "5_2",
      "unit": "V",
      "decimals": 2
    },
    // ... 更多映射规则 ...
  },
  "logData": [ // 日志数据记录列表
    {
      "sn": "SN1234567890", // 序列号
      "detailedItems": [ // 该记录的详细测试项列表
        {
          "itemName": "电池电压",
          "actualValue": "3.752"
        },
        {
          "itemName": "充电电流",
          "actualValue": "0.501"
        }
        // ... 更多测试项 ...
      ]
    },
    {
      "sn": "SN9876543210",
      "detailedItems": [
        {
          "itemName": "电池电压",
          "actualValue": "3.801"
        },
         {
          "itemName": "充电电流",
          "actualValue": "0.495"
        }
      ]
    }
    // ... 更多日志记录 ...
  ]
}```

### 3. 响应体

*   **成功 (HTTP Status 200 OK):**
    *   **Content-Type:**
        *   `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet` (对于 `single-sheet` 和 `multi-sheet` 模式)
        *   `application/zip` (对于 `zip-files` 模式)
    *   **Headers:** 包含 `Content-Disposition` 头，提供生成的文件名（通常包含 SN 或时间戳）。
    *   **Body:** 生成的 Excel 文件或 ZIP 文件的原始字节流。前端接收此流并触发下载。

*   **失败 (HTTP Status 4xx 或 5xx):**
    *   **Body:** 可能包含简单的错误信息字符串。请检查后端日志获取详细信息。

## 前端集成

前端应用程序（`reportGenerator.js`）需要修改其导出逻辑，以：

1.  监听导出按钮点击事件。
2.  收集当前的 Excel 模板文件对象、选中的日志数据 (`logData.records`) 和映射规则 (`currentMapping`)。
3.  根据后端 API 要求，构建 `multipart/form-data` 请求体。
4.  使用 `fetch` 或其他 HTTP 客户端向后端 `/api/reports/generate` 端点发起 POST 请求。
5.  处理后端的响应，特别是接收文件字节流，并根据响应头中的 `Content-Disposition` 提供的文件名，使用前端的下载逻辑（如 `downloadBlob` 函数）触发文件下载。
6.  在请求发送和处理过程中，更新 UI 状态（如显示“导出中”）。

## 限制与注意事项

*   **Sheet 复制的局限性:** `copySheetContent` 方法尽力复制 Sheet 内容，但无法保证完美复制所有 Excel 特性，特别是复杂的绘图对象（如图表、SmartArt、组合图形）、复杂的条件格式、数据验证、宏（VBA）、ActiveX 控件、评论和超链接等。这些元素可能不会被复制或复制后效果不完全一致。
*   **样式和颜色的精确性:** 虽然引入了样式映射，但由于源和目标 Workbook 的内部结构、主题、自定义调色板可能不同，某些颜色（尤其是主题颜色或非标准索引颜色）在复制后可能仍然存在细微差异。
*   **性能:** 对于包含大量数据或复杂样式/图片的超大模板文件和海量日志记录，生成报告可能需要较长时间并消耗较多内存。
*   **错误处理:** 当前实现的错误处理主要是打印日志和返回通用 HTTP 状态码。更健壮的应用需要细化后端异常处理和前端错误提示。

## 潜在的未来增强

*   **更完善的 Sheet 复制:** 深入研究 POI API 或其他库，实现对更多类型绘图对象（如文本框、简单形状）甚至图表的复制。
*   **性能优化:** 对于大规模数据，可以考虑流式处理或使用更高效的库。
*   **错误日志记录和监控:** 集成日志框架（如 Logback, Log4j2）和监控工具。
*   **配置化:** 将端口、API 路径等信息放入配置文件。
*   **更多导出选项:** 例如 PDF 导出（需要额外的库和配置）。
*   **异步处理:** 对于耗时长的报告生成任务，可以考虑异步处理，避免阻塞 HTTP 请求。
