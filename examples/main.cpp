/*
 * Project: xls2img
 * Repository: https://github.com/capp-adocia/xls2img
 * Author: SiLan (https://github.com/capp-adocia)
 *
 * Copyright (c) 2026 SiLan
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

#define _CRT_SECURE_NO_WARNINGS
#include <iostream>
#include <fstream>
#include <memory>
#include <windows.h>
#include <string>
#include <xls2img.h>

static bool read_file(const wchar_t* filepath, std::unique_ptr<unsigned char[]>& buffer, size_t& size) {
    FILE* fp = _wfopen(filepath, L"rb");
    if (!fp)
    {
        std::wcerr << L"Error: Failed to open file: " << filepath << std::endl;
        return false;
    }

    fseek(fp, 0, SEEK_END);
    long file_size = ftell(fp);
    fseek(fp, 0, SEEK_SET);

    if (file_size <= 0)
    {
        fclose(fp);
        std::wcerr << L"Error: File size is zero or invalid" << std::endl;
        return false;
    }

    buffer = std::make_unique<unsigned char[]>(file_size);
    size_t bytes_read = fread(buffer.get(), 1, file_size, fp);
    fclose(fp);

    if (bytes_read != static_cast<size_t>(file_size))
    {
        std::wcerr << L"Error: File read incomplete" << std::endl;
        return false;
    }

    size = bytes_read;
    return true;
}

static bool save_image(const char* filename, const void* data, size_t size)
{
    std::ofstream file(filename, std::ios::binary);
    if (!file.is_open())
    {
        std::cerr << "Error: Cannot create file " << filename << std::endl;
        return false;
    }

    file.write(static_cast<const char*>(data), size);
    return file.good();
}

int main()
{
    // Use default test file
    const wchar_t* filepath = L"./test.xls";

    // Read file
    std::unique_ptr<unsigned char[]> file_buffer;
    size_t file_size = 0;

    if (!read_file(filepath, file_buffer, file_size))
        return -1;

    // Initialize xls2img reader
    XLS2IMG_READER* reader = nullptr;
    int ret = xls2img_open(&reader, file_buffer.get(), file_size);
    if (ret != XLS2IMG_SUCCESS)
    {
        std::cerr << "Initialization failed: " << xls2img_strerror(ret) << std::endl;
        return -1;
    }

    // Extract workbook stream
    void* workbook_data = nullptr;
    size_t workbook_size = 0;

    ret = xls2img_get_workbook(reader, &workbook_data, &workbook_size);
    if (ret != XLS2IMG_SUCCESS)
    {
        std::cerr << "Failed to extract workbook: " << xls2img_strerror(ret) << std::endl;
        xls2img_close(reader);
        return -1;
    }

    // Extract images
    XLS2IMG_RESULT images = { nullptr, 0 };
    ret = xls2img_extract_images(workbook_data, workbook_size, &images);

    // Process results
    if (ret > 0)
    {
        std::cout << "Extracted " << images.count << " images" << std::endl;

        for (int i = 0; i < images.count; ++i)
        {
            const XLS2IMG_IMAGE& img = images.images[i];
            const char* format_str = (img.format == XLS2IMG_PNG) ? "PNG" : "JPEG";

            std::cout << "Image " << (i + 1) << ": Format=" << format_str << ", Size=" << img.size << " bytes" << std::endl;

            std::string filename = "image_" + std::to_string(i + 1) + ((img.format == XLS2IMG_PNG) ? ".png" : ".jpg");

            if (save_image(filename.c_str(), img.data, img.size))
                std::cout << "  -> Saved to: " << filename << std::endl;
            else
                std::cerr << "  -> Failed to save: " << filename << std::endl;
        }

        xls2img_free_result(&images);
    }
    else
        std::cerr << "Failed to extract images: " << xls2img_strerror(ret) << std::endl;

    // Cleanup
    xls2img_free_result(&images);
    xls2img_free_workbook_data(workbook_data);
    xls2img_close(reader);

    std::cout << "Image extraction completed." << std::endl;
    return 0;
}