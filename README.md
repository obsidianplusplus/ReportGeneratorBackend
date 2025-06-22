# Report Generator Backend 📊

## 项目描述 ✨

这是一个基于 Spring Boot 和 Apache POI 构建的后端服务，专门为前端的日志分析仪表盘项目提供报告生成功能。它被设计为一个“哑处理器”，不包含复杂的业务逻辑，完全依赖于前端提供的 Excel 模板文件 📄、解析后的日志数据 📝 以及详细的映射规则 🗺️ 来动态生成 Excel 报告。

本项目的主要目标是接管报告生成中与文件操作相关的复杂任务，例如加载 Excel 模板、根据指令填充数据、复制 Sheet 结构（包括样式 ✨ 和图片 🖼️ 等）以及处理不同的导出格式（单表, 多 Sheet, ZIP 包 🗄️），从而减轻前端的负担并利用 Java 在文件处理方面的优势 💪。

## 主要功能 🚀

*   接收前端通过 HTTP POST 请求上传的 Excel 模板文件 📤。
*   接收前端发送的 JSON 数据 📨，包含待填充的日志记录列表和映射规则。
*   支持多种导出模式 💾：
    *   **Single Sheet:** 将所有选中的日志记录的数据，根据映射规则，按列偏移填充到模板的第一个 Sheet 中。
    *   **Multi-Sheet:** 为每一条选中的日志记录，创建一个新的 Sheet（基于模板第一个 Sheet 的副本），并填充该记录的数据。新 Sheet 的名称通常基于记录的 SN。所有 Sheet 合并在一个 Excel 文件中 📚。
    *   **ZIP Files:** 为每一条选中的日志记录，生成一个独立的 Excel 文件，然后将所有生成的 Excel 文件压缩成一个 ZIP 包 📦。
*   复制模板 Sheet 的内容，包括：
    *   单元格值和类型 📝。
    *   大部分单元格样式 ✨（通过样式映射机制，以提高颜色 🎨 和字体的准确性）。
    *   合并单元格区域 🔠。
    *   列宽 ↔️。
    *   基础图片 🖼️（限于 XSSFClientAnchor 类型的图片）。
*   将生成的 Excel 文件或 ZIP 包作为 HTTP 响应流返回给前端 📥。

## 技术栈 🛠️

*   **后端框架:** Spring Boot 3.x 🍃
*   **Java 版本:** Java 17+ ☕
*   **Excel 处理:** Apache POI 5.x 📊
*   **构建工具:** Maven ⚙️
*   **依赖管理:** Spring Boot Starter Parent
*   **开发便利:** Lombok (可选，用于简化 JavaBean 代码)


*   JDK 17 或更高版本
*   Maven 3.x
*   IntelliJ IDEA, Eclipse 或其他兼容 Spring Boot 的 IDE (推荐 IntelliJ IDEA 😎)

## 快速开始 ▶️
