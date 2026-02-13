# xls2img
Chinese | [English](./README.md)

`xls2img` 是一个用 C 语言编写的库，旨在从 Microsoft Excel 97-2003 格式的 `.xls` 文件中提取嵌入的图片（如 JPEG、PNG）。

## 为什么需要它
* libxls不支持图片
* 调用Python/PHP/Java方案太重
* 商业库有许可风险

## 实现

### 提取思路

1. **获取 WorkBook 流**
   *   根据微软相关文档，像 JPG、PNG 这样的文件都存储在 WorkBook 流中。因此，首先需要解析 XLS 的文件格式。了解 XLS 格式非常困难，感谢微软的 [compoundfilereader](https://github.com/microsoft/compoundfilereader) 仓库，它极大地简化了这一流程，使我们能够从 XLS 中解析出 WorkBook 流。

2. **从 WorkBook 流中查找 `MsoDrawingGroup`**
   *   WorkBook 流由多个 BIFF8 结构组成，每个 BIFF8 都包含记录类型、记录大小和记录数据。需要提取的图片数据就对应记录类型为 `MsoDrawingGroup` 的结构。该结构中存储了嵌入在 XLS 中的图像数据。通过遍历每个 BIFF8 结构来获取所有的 `MsoDrawingGroup` 数据。
   *   由于单个 BIFF8 的存储有上限，一个图像的数据可能会被分割存储在多个 BIFF8 中。幸运的是，每当图像数据被分割时，除了第一个 BIFF8 外，后续的数据块会以 `recordType = BIFF8_CONTINUE` 作为标志，这使得我们可以将这些分离的数据块重新组合。

3.  **从 `MsoDrawingGroup` 中解析图像数据**
    *   `MsoDrawingGroup` 记录不仅包含图片的原始数据，还混杂着描述图形布局、位置等用途的元信息。因此，不能将其整个数据块直接写入文件。需要进一步解析以提取出纯净的图像数据。虽然微软提供了相关的官方文档，但内容庞杂，直接依据文档解析图片数据流程复杂且耗时。
    *   通过对 `MsoDrawingGroup` 二进制数据的分析发现，其中绝大部分内容就是原始的图片数据，并且这些图片数据是按顺序、完整地存储的。基于此观察，采用一种**基于文件签名的顺序流分割算法**来解析，是一种高效且可行的方案。

4.  **基于文件签名的顺序流分割算法**
    *   `xls2img` 专门针对 PNG 和 JPEG 两种图像格式进行处理，其核心实现思路如下：
        1.  **PNG:** 遍历 `MsoDrawingGroup` 的数据流，精确匹配 PNG 文件特有的 8 字节文件头签名（`89 50 4E 47 0D 0A 1A 0A`）。该签名具有极高的辨识度，误判率非常低。定位到文件头后，再严格按照 PNG 格式规范，查找其对应的 `IEND` 结束块，从而确定图片数据的完整边界。
        2.  **JPEG:** 同样遍历数据流，查找 JPEG 文件的起始标识。常见的标识包括 `JFIF` (应用段 APP0) 和 `Exif` (应用段 APP1) 的头部，这些签名同样具有良好的区分性。对于结束标志，JPEG 文件的标准结尾是 `0xFF 0xD9`。然而，`0xFF 0xD9` 这个字节序列有时也可能出现在图像的实际像素数据（内容）中。如果采用从前往后顺序扫描的方式查找结束符，极易捕获到这些“假”的结束符，从而将一张完整的图片错误地分割成多段。为规避此问题，本库采用了**从后往前回溯**的策略来定位 `0xFF 0xD9`。具体的触发时机是：当成功定位到下一个有效的图像文件头时，便从此新文件头的位置开始，向前回溯，查找上一张图片的真实结束符。
    *   通过以上方法，即可完成对这两种主流格式图片的精准提取。关于性能表现，请参见下方测试。

## 命令行工具

除了作为库使用，`xls2img` 还提供了一个方便的命令行工具，用于直接提取图片。

**下载:** 请前往项目的 [Releases](https://github.com/capp-adocia/xls2img/releases) 页面，下载名为 `xls2img_tool.zip` 的压缩包。

## 性能和准确度

**性能测试**（在 Release 模式下解析 Workbook 流并提取 XLS 中的所有图片）：

| 文件大小 | 图片数量 | 实际提取数量 | 状态 | 详情 |
| :--- | :--- | :--- | :--- | :--- |
| 6.5 MB | 6 张 | 6 张 | 提取图像与源图像完全一致 | ![Small File Extraction Result](./doc/release_extrator_small.png) |
| 47.5 MB | 20 张 | 20 张 | 提取图像与源图像完全一致 | ![Big File Extraction Result](./doc/release_extrator_big.png) |

**总结：** 经过多轮测试,性能还算可以，尚未发现特殊情况下无法正确提取的情况。

## 安装

### 从源码构建

1.  **克隆仓库:**
    ```bash
    git clone https://github.com/capp-adocia/xls2img.git
    cd xls2img
    ```

2.  **构建:**
    *   **Windows (MSVC):**
        ```bash
        cmake -B build
        cmake --build build --config Release
        ```
        生成的文件位于 `build/` 目录中。
    *   **Linux (暂未支持)**

3.  **安装:**
    *   构建过程生成了 `.lib` 和 `.dll` 文件。您可以将这些库文件以及 `xls2img.h` 头文件复制到您的项目或系统的库目录中。

## 快速开始

以下是一个简单的 C 语言示例，展示如何使用 `xls2img` 库，该案例来源于 `examples/main.c`：

```c
int main()
{
    // Use default test file
    const wchar_t* filepath = L"./test.xls";

    // Read file
    unsigned char* file_buffer = NULL;
    size_t file_size = 0;

    if (!read_file(filepath, &file_buffer, &file_size))
    {
        return -1;
    }

    // Initialize xls2img reader
    XLS2IMG_READER* reader = NULL;
    int ret = xls2img_open(&reader, file_buffer, file_size);
    if (ret != XLS2IMG_SUCCESS)
    {
        fprintf(stderr, "Initialization failed: %s\n", xls2img_strerror(ret));
        free(file_buffer);
        return -1;
    }

    // Extract workbook stream
    void* workbook_data = NULL;
    size_t workbook_size = 0;

    ret = xls2img_get_workbook(reader, &workbook_data, &workbook_size);
    if (ret != XLS2IMG_SUCCESS)
    {
        fprintf(stderr, "Failed to extract workbook: %s\n", xls2img_strerror(ret));
        xls2img_close(reader);
        free(file_buffer);
        return -1;
    }

    // Extract images
    XLS2IMG_RESULT images = { NULL, 0 };
    ret = xls2img_extract_images(workbook_data, workbook_size, &images);

    // Process results
    if (ret > 0)
    {
        printf("Extracted %d images\n", images.count);

        // Process each image
        for (int i = 0; i < images.count; i++)
        {
            const XLS2IMG_IMAGE* img = &images.images[i];
            const char* format_str = (img->format == XLS2IMG_PNG) ? "PNG" : "JPEG";

            printf("Image %d: Format=%s, Size=%zu bytes\n", i + 1, format_str, img->size);
        }
    }
    else
        fprintf(stderr, "Failed to extract images: %s\n", xls2img_strerror(ret));

    // Cleanup
    free(file_buffer);
    xls2img_free_result(&images);
    xls2img_free_workbook_data(workbook_data);
    xls2img_close(reader);

    printf("Image extraction completed.\n");
    return 0;
}
```

## API 文档

详细的 API 文档可以在 `xls2img.h` 头文件中找到，其中包含了每个函数、枚举和结构体的详细说明。

## 其他

### 构建系统

本项目使用 [CMake](https://cmake.org/) 作为构建系统。请确保您的系统已安装 CMake。

### 依赖项

*   **C Standard Library:** 项目依赖标准 C 库。

### 架构概览

*   `xls2img.h`: 公共 API 接口定义。
*   `xls2img_reader.c`: XLS 文件解析和读取逻辑的核心实现。
*   `xls2img_images.c`: 图片提取和处理逻辑的核心实现。

## 许可证

本项目采用 [MIT License](./LICENSE)。详见 `LICENSE` 文件。

## 致谢

*   再次感谢 [microsoft](https://github.com/microsoft/compoundfilereader) 的 `compoundfilereader` 项目及其相关工具和文档，为理解 XLS 文件格式提供了参考。
