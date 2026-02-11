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
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <io.h>
#include <fcntl.h>
#include <wchar.h>
#include <windows.h>
#include <shellapi.h>
#include <xls2img.h>

static int read_file(const wchar_t* filepath, unsigned char** buffer, size_t* size)
{
    FILE* fp = _wfopen(filepath, L"rb");
    if (!fp)
    {
        fwprintf(stderr, L"Error: Failed to open file: %ls\n", filepath);
        return 0;
    }

    fseek(fp, 0, SEEK_END);
    long file_size = ftell(fp);
    fseek(fp, 0, SEEK_SET);

    if (file_size <= 0)
    {
        fclose(fp);
        fwprintf(stderr, L"Error: File size is zero or invalid\n");
        return 0;
    }

    *buffer = (unsigned char*)malloc(file_size);
    if (!*buffer)
    {
        fclose(fp);
        fwprintf(stderr, L"Error: Memory allocation failed\n");
        return 0;
    }

    size_t bytes_read = fread(*buffer, 1, file_size, fp);
    fclose(fp);

    if (bytes_read != (size_t)file_size)
    {
        free(*buffer);
        *buffer = NULL;
        fwprintf(stderr, L"Error: File read incomplete\n");
        return 0;
    }

    *size = bytes_read;
    return 1;
}

// Updated save_image function using Windows API
static int save_image(const wchar_t* filename, const void* data, size_t size)
{
    HANDLE hFile = CreateFileW(filename, GENERIC_WRITE, 0, NULL,
        CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    if (hFile == INVALID_HANDLE_VALUE)
    {
        fwprintf(stderr, L"Error: Cannot create file %ls\n", filename);
        return 0;
    }

    DWORD written = 0;
    BOOL result = WriteFile(hFile, data, (DWORD)size, &written, NULL);
    CloseHandle(hFile);

    return (result && written == size);
}

int wmain(int argc, wchar_t* argv[])
{
    _setmode(_fileno(stdout), _O_U16TEXT);
    _setmode(_fileno(stderr), _O_U16TEXT);

    if (argc < 2)
    {
        fwprintf(stderr, L"Usage: \n./xls2img_tool.exe <input.xls> [output_dir]\n");
        return -1;
    }

    const wchar_t* input_path = argv[1];
    const wchar_t* output_path = (argc > 2) ? argv[2] : L"./";

    // Read file
    unsigned char* file_buffer = NULL;
    size_t file_size = 0;

    if (!read_file(input_path, &file_buffer, &file_size))
    {
        return -1;
    }

    // Initialize xls2img reader
    XLS2IMG_READER* reader = NULL;
    int ret = xls2img_open(&reader, file_buffer, file_size);
    if (ret != XLS2IMG_SUCCESS)
    {
        fwprintf(stderr, L"Initialization failed: %hs\n", xls2img_strerror(ret));
        free(file_buffer);
        return -1;
    }

    // Extract workbook stream
    void* workbook_data = NULL;
    size_t workbook_size = 0;

    ret = xls2img_get_workbook(reader, &workbook_data, &workbook_size);
    if (ret != XLS2IMG_SUCCESS)
    {
        fwprintf(stderr, L"Failed to extract workbook: %hs\n", xls2img_strerror(ret));
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
        wprintf(L"Extracted %d images\n", images.count);

        // Process each image
        for (int i = 0; i < images.count; i++)
        {
            const XLS2IMG_IMAGE* img = &images.images[i];
            const wchar_t* format_str = (img->format == XLS2IMG_PNG) ? L"png" : L"jpg";

            // Prepare filename with wide characters
            wchar_t w_filename[MAX_PATH];
            swprintf_s(w_filename, MAX_PATH, L"%lsimage_%d.%ls", output_path, i + 1, format_str);

            wprintf(L"Image %d: Format=%ls, Size=%zu bytes\n", i + 1, format_str, img->size);

            // Save image using Windows API
            if (save_image(w_filename, img->data, img->size))
                wprintf(L"  -> Saved to: %ls\n", w_filename);
            else
                fwprintf(stderr, L"  -> Failed to save: %ls\n", w_filename);
        }

        // Free images
        xls2img_free_result(&images);
    }
    else
        fwprintf(stderr, L"Failed to extract images: %hs\n", xls2img_strerror(ret));

    // Cleanup
    xls2img_free_result(&images);
    xls2img_free_workbook_data(workbook_data);
    xls2img_close(reader);

    wprintf(L"Image extraction completed.\n");
    return 0;
}