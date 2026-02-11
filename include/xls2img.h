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

#ifndef XLS2IMG_H
#define XLS2IMG_H

#include <stdint.h>
#include <stddef.h>

#ifdef __cplusplus
extern "C" {
#endif

#if defined(_WIN32)
    #if defined(XLS2IMG_BUILD_DLL)
        #define XLS2IMG_API __declspec(dllexport)
    #elif defined(XLS2IMG_USE_DLL)
        #define XLS2IMG_API __declspec(dllimport)
    #else
        #define XLS2IMG_API
    #endif
#else
    #if defined(XLS2IMG_BUILD_SHARED)
        #define XLS2IMG_API __attribute__((visibility("default")))
    #else
        #define XLS2IMG_API
    #endif
#endif

// err code
#define XLS2IMG_SUCCESS 0
#define XLS2IMG_ERROR_WRONG_FORMAT -1
#define XLS2IMG_ERROR_FILE_CORRUPTED -2
#define XLS2IMG_ERROR_INVALID_ARGUMENT -3
#define XLS2IMG_ERROR_NO_WORKBOOK -4
#define XLS2IMG_ERROR_NO_IMAGES -5
#define XLS2IMG_ERROR_OUT_OF_MEMORY -6

    /**
     * @brief Image format enumeration
     */
    typedef enum {
        XLS2IMG_UNKNOWN = 0,  /* Unknown format */
        XLS2IMG_PNG = 1,      /* PNG format */
        XLS2IMG_JPG = 2,      /* JPG/JPEG format */
    } XLS2IMG_FORMAT;

    /**
     * @brief Image information structure
     */
    typedef struct {
        XLS2IMG_FORMAT format;  /* Image format */
        size_t size;            /* Image data size */
        void* data;             /* Image data pointer */
    } XLS2IMG_IMAGE;

    /**
     * @brief Image extraction result struct
     */
    typedef struct {
        XLS2IMG_IMAGE* images;  /* Image array pointer */
        int count;              /* Number of images */
    } XLS2IMG_RESULT;

    /**
     * @brief XLS file reader context
     */
    typedef struct XLS2IMG_READER XLS2IMG_READER;

    /**
     * @brief Get error message from error code
     * @param[in] error_code Error code returned by xls2img functions
     * @return Pointer to error message string
     */
    XLS2IMG_API const char* xls2img_strerror(int error_code);

    /**
     * @brief Create a reader for XLS file data
     * @param[out] reader Returns the created reader pointer
     * @param[in] buffer XLS file data buffer
     * @param[in] len buffer size
     * @return XLS2IMG_SUCCESS on success, error code on failure
     */
    XLS2IMG_API int xls2img_open(XLS2IMG_READER** reader, const void* buffer, size_t len);

    /**
     * @brief Close and free the reader
     * @param[in] reader The reader pointer to deallocate
     */
    XLS2IMG_API void xls2img_close(XLS2IMG_READER* reader);

    /**
     * @brief Extract workbook stream from XLS file
     * @param[in] reader XLS2IMG reader
     * @param[out] data Output parameter, returns workbook data pointer
     * @param[out] size Output parameter, returns workbook data size
     * @return XLS2IMG_SUCCESS on success, error code on failure
     */
    XLS2IMG_API int xls2img_get_workbook(XLS2IMG_READER* reader, void** data, size_t* size);

    /** 
     *  @brief Frees the workbook data buffer allocated by xls2img_get_workbook.
     *  @param[in] data The pointer returned by xls2img_get_workbook.
     */
    XLS2IMG_API void xls2img_free_workbook_data(void* data);

    /**
     * @brief Extract images from workbook data
     * @param[in] workbook_data workbook stream data pointer
     * @param[in] workbook_size workbook stream size
     * @param[out] result Output parameter, returns extracted image information
     * @return Number of images extracted on success (>=1), error code on failure (<=0)
     */
    XLS2IMG_API int xls2img_extract_images(const void* workbook_data, size_t workbook_size, XLS2IMG_RESULT* result);

    /**
     * @brief Free image extraction results
     * @param[in] result The extracted image result to release
     */
    XLS2IMG_API void xls2img_free_result(XLS2IMG_RESULT* result);


#ifdef __cplusplus
}
#endif

#endif /* XLS2IMG_H */