# extract_word_images
自动解析.docx文件结构，快速提取文档内嵌的所有图片（支持PNG/JPG/EMF等格式），无需手动解压操作。
Word图片提取工具 v1.0 —— 开发者必备的文档处理利器
by drake | 2025-02-17  | mail:drake816@163.com
核心功能
一键提取Word图片
自动解析.docx文件结构，快速提取文档内嵌的所有图片（支持PNG/JPG/EMF等格式），无需手动解压操作。
EMF智能转PNG
自动检测并转换Windows矢量图格式（.emf）为通用PNG格式，解决跨平台显示兼容性问题。
可视化进度追踪
实时进度条+日志面板，清晰展示解压、转换、保存全过程，处理状态一目了然。


技术亮点

轻量级架构：基于Python Tkinter开发，仅需标准库（zipfile/shutil）和Pillow图像库
异常处理机制：自动清理临时文件，错误弹窗提示保障程序健壮性
跨平台支持：兼容Windows/macOS/Linux系统（需Python 3.6+环境）


三步极简操作
1.双击可执行应用

2.选择文档：点击"选择文件"按钮导入.docx文件

3.设置路径：指定图片保存目录（默认创建"temp_extract"临时文件夹）

4.开始提取：点击绿色按钮启动，自动完成解压→转换→清理全流程



⚡ 适用场景

批量导出技术文档/论文中的图表
快速提取PPT转存Word的矢量图素材
自动化测试中的文档内容验证


立即体验
 [项目仓库地址]https://github.com/Dreke666/extract_word_images

（建议Python 3.8+环境，依赖安装：pip install pillow）

开发者注：本工具遵循MIT协议开源，欢迎贡献代码或提交Issue优化建议。
