/*
 * Project: xls2img
 * Repository: https://github.com/capp-adocia/Xls2Img
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

#include "xls2img.h"
#include <string.h>
#include <stdlib.h>
#include <stdio.h>

// BIFF8 Record Type
enum {
    BIFF8_EOF_RECORD = 0x000A,
    BIFF8_MsoDrawingGroup = 0x00EB,
    BIFF8_CONTINUE = 0x003C,
};

// Image processing
static XLS2IMG_FORMAT xls2img_identify_format(const uint8_t* data)
{
    // PNG Signature
    if (data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47 &&
        data[4] == 0x0D && data[5] == 0x0A && data[6] == 0x1A && data[7] == 0x0A)
        return XLS2IMG_PNG;

    // JPG Signature
    if (data[0] == 0xFF && data[1] == 0xD8)
        return XLS2IMG_JPG;

    return XLS2IMG_UNKNOWN;
}

// quickly find the next possible image header within a data block.
static const uint8_t* xls2img_find_next_header(const uint8_t* data, size_t size)
{
    // Subtract the maximum number of bytes that might need to be checked to prevent out-of-bounds access.
    // PNG requires 8 bytes, JPEG (SOI + APPn + ID string) needs at least 2 (SOI) + 2 (Marker) + 2 (Len) + 4 (ID) = 10 bytes.
    for (size_t i = 0; i < size - 9; i++)
    {
        uint8_t firstByte = data[i];
        uint8_t secondByte = data[i + 1];

        // PNG Signature: 89 50 4E 47 0D 0A 1A 0A
        if (firstByte == 0x89 && secondByte == 0x50 &&
            data[i + 2] == 0x4E && data[i + 3] == 0x47 &&
            data[i + 4] == 0x0D && data[i + 5] == 0x0A &&
            data[i + 6] == 0x1A && data[i + 7] == 0x0A) {
            return data + i;
        }

        // Check for SOI marker
        if (firstByte == 0xFF && secondByte == 0xD8)
        {
            if (i + 3 < size)
            {
                uint8_t app_marker_high = data[i + 2];
                uint8_t app_marker_low = data[i + 3];

                // JFIF
                if (app_marker_high == 0xFF && app_marker_low == 0xE0)
                {
                    size_t app0_length_offset = i + 4;
                    if (app0_length_offset + 1 < size)
                    {
                        uint16_t app0_length = ((uint16_t)data[app0_length_offset] << 8) | data[app0_length_offset + 1];

                        size_t identifier_start_offset = app0_length_offset + 2;
                        if (identifier_start_offset + 4 <= size)
                        {
                            if (data[identifier_start_offset] == 'J' &&
                                data[identifier_start_offset + 1] == 'F' &&
                                data[identifier_start_offset + 2] == 'I' &&
                                data[identifier_start_offset + 3] == 'F') {
                                return data + i;
                            }
                        }
                    }
                }
                // Exif
                else if (app_marker_high == 0xFF && app_marker_low == 0xE1)
                {
                    size_t app1_length_offset = i + 4;
                    if (app1_length_offset + 1 < size)
                    {
                        uint16_t app1_length = ((uint16_t)data[app1_length_offset] << 8) | data[app1_length_offset + 1];

                        size_t identifier_start_offset = app1_length_offset + 2;
                        if (identifier_start_offset + 4 <= size)
                        {
                            if (data[identifier_start_offset] == 'E' &&
                                data[identifier_start_offset + 1] == 'x' &&
                                data[identifier_start_offset + 2] == 'i' &&
                                data[identifier_start_offset + 3] == 'f') {
                                return data + i;
                            }
                        }
                    }
                }
            }
        }
    }
    return NULL;
}

static int xls2img_find_png_end(const uint8_t* data, size_t size)
{
    if (size < 8) return -1;

    const uint8_t* p = data;
    const uint8_t* end = data + size;

    if (!(p[0] == 0x89 && p[1] == 0x50 && p[2] == 0x4E && p[3] == 0x47))
        return -1;

    p += 8;

    while (p < end)
    {
        if (end - p < 8) return -1;

        uint32_t chunk_len = (p[0] << 24) | (p[1] << 16) | (p[2] << 8) | p[3];
        const uint8_t* chunk_type = p + 4;

        if (chunk_type[0] == 'I' && chunk_type[1] == 'E' &&
            chunk_type[2] == 'N' && chunk_type[3] == 'D')
            return (int)((p + 12) - data); // 4(Len) + 4(Type) + 4(CRC)

        p += 12 + chunk_len;
        if (p > end) return -1;
    }
    return -1;
}

static int xls2img_find_jpg_end(const uint8_t* start_search_from, const uint8_t* actual_start_of_last_jpg)
{
    const uint8_t* p = start_search_from - 2;
    const uint8_t* limit = actual_start_of_last_jpg;

    while (p >= limit)
    {
        if (p[0] == 0xFF && p[1] == 0xD9)
            return (int)((p + 2) - actual_start_of_last_jpg); // Returns the size relative to the previous JPEG starting position
        p--;
    }
    return -1;
}

// Buffer collection structure
typedef struct {
    uint8_t* data;
    size_t size;
    size_t capacity;
} BufferCollector;

static void buffer_collector_init(BufferCollector* collector)
{
    collector->data = NULL;
    collector->size = 0;
    collector->capacity = 0;
}

static int buffer_collector_append(BufferCollector* collector, const uint8_t* data, size_t size)
{
    if (collector->size + size > collector->capacity)
    {
        // grow in larger chunks and avoid frequent realloc
        size_t new_capacity = collector->capacity * 2;
        if (new_capacity < collector->size + size)
            new_capacity = collector->size + size;

        // add some extra space to avoid realloc immediately next time
        new_capacity = new_capacity * 3 / 2;

        uint8_t* new_data = (uint8_t*)realloc(collector->data, new_capacity);
        if (!new_data)
            return 0;

        collector->data = new_data;
        collector->capacity = new_capacity;
    }

    memcpy(collector->data + collector->size, data, size);
    collector->size += size;

    return 1;
}

static void buffer_collector_free(BufferCollector* collector)
{
    if (collector->data)
    {
        free(collector->data);
        collector->data = NULL;
    }
    collector->size = 0;
    collector->capacity = 0;
}

// Helper function to add an image to the result array.
static int add_image_to_result(XLS2IMG_IMAGE** images, int* capacity, int* count, const uint8_t* data, size_t size, XLS2IMG_FORMAT format)
{
    if (*count >= *capacity)
    {
        int new_capacity = *capacity * 2;
        XLS2IMG_IMAGE* new_images = realloc(*images, new_capacity * sizeof(XLS2IMG_IMAGE));
        if (!new_images) return 0;
        *images = new_images;
        *capacity = new_capacity;
    }

    void* image_data = malloc(size);
    if (!image_data) return 0;

    memcpy(image_data, data, size);
    (*images)[*count].format = format;
    (*images)[*count].size = size;
    (*images)[*count].data = image_data;
    (*count)++;
    return 1;
}

// Helper function to finalize and add the last tracked image if valid.
static void process_last_image_if_any(const uint8_t** last_img_start, XLS2IMG_FORMAT* last_img_fmt, XLS2IMG_IMAGE** images,
    int* capacity, int* count, const uint8_t* current_block_ptr_or_end, const uint8_t* last_img_block_start)
{
    if (*last_img_start != NULL)
    {
        int img_size = -1;
        if (*last_img_fmt == XLS2IMG_PNG)
            img_size = xls2img_find_png_end(*last_img_start, current_block_ptr_or_end - *last_img_start);
        else if (*last_img_fmt == XLS2IMG_JPG)
            img_size = xls2img_find_jpg_end(current_block_ptr_or_end, *last_img_start);

        if (img_size > 0)
            add_image_to_result(images, capacity, count, *last_img_start, img_size, *last_img_fmt);

        // Reset the tracking variables after attempting to process.
        *last_img_start = NULL;
        *last_img_fmt = XLS2IMG_UNKNOWN;
    }
}

int xls2img_extract_images(const void* workbook_data, size_t workbook_size, XLS2IMG_RESULT* result)
{
    if (!workbook_data || workbook_size == 0 || !result)
        return XLS2IMG_ERROR_INVALID_ARGUMENT;

    result->images = NULL;
    result->count = 0;

    const uint8_t* ptr = (const uint8_t*)workbook_data;
    const uint8_t* end_ptr = ptr + workbook_size;

    int capacity = 16;
    int count = 0;
    XLS2IMG_IMAGE* images = (XLS2IMG_IMAGE*)malloc(capacity * sizeof(XLS2IMG_IMAGE));
    if (!images)
        return XLS2IMG_ERROR_INVALID_ARGUMENT;

    BufferCollector mso_collector;
    buffer_collector_init(&mso_collector);
    int collecting_mso = 0;

    // Used to record the starting position and format of the previous image
    const uint8_t* last_img_start = NULL;
    XLS2IMG_FORMAT last_img_fmt = XLS2IMG_UNKNOWN;

    while (ptr < end_ptr)
    {
        if (end_ptr - ptr < 4) break;

        uint16_t recordType = (ptr[1] << 8) | ptr[0];
        uint16_t recordSize = (ptr[3] << 8) | ptr[2];
        ptr += 4;

        if (recordType == BIFF8_MsoDrawingGroup)
        {
            collecting_mso = 1;

            // reset the collector to pre-allocate enough space
            buffer_collector_free(&mso_collector);
            mso_collector.data = (uint8_t*)malloc(recordSize * 2);
            if (mso_collector.data)
            {
                mso_collector.capacity = recordSize * 2;
                memcpy(mso_collector.data, ptr, recordSize);
                mso_collector.size = recordSize;
            }
        }
        else if (collecting_mso && recordType == BIFF8_CONTINUE)
        {
            // append to the MsoDrawingGroup data
            if (!buffer_collector_append(&mso_collector, ptr, recordSize))
            {
                buffer_collector_free(&mso_collector);
                free(images);
                return XLS2IMG_ERROR_INVALID_ARGUMENT;
            }
        }
        else if (collecting_mso)
        {
            // this means that the MsoDrawingGroup chain is complete and that it is time to collect the image data
            // there should be only one MsoDrawingGroup in the xls, so it will be executed only once
            if (mso_collector.data && mso_collector.size > 0)
            {
                const uint8_t* block_ptr = mso_collector.data;
                const uint8_t* block_end = block_ptr + mso_collector.size;

                while (block_ptr < block_end)
                {
                    const uint8_t* next_header = xls2img_find_next_header(block_ptr, block_end - block_ptr);

                    // Process the previous image before moving the pointer.
                    if (last_img_start != NULL && next_header != NULL)
                        process_last_image_if_any(&last_img_start, &last_img_fmt, &images, &capacity, &count, next_header, last_img_start);

                    if (!next_header)
                        break;

                    // Record the current found image header as the next candidate to be processed.
                    block_ptr = next_header;
                    XLS2IMG_FORMAT fmt = xls2img_identify_format(block_ptr);
                    if (fmt == XLS2IMG_PNG || fmt == XLS2IMG_JPG)
                    {
                        last_img_start = block_ptr;
                        last_img_fmt = fmt;
                    }
                    block_ptr++;
                }

                // After the loop, check if the last image was not processed (e.g., at the end of the block).
                process_last_image_if_any(&last_img_start, &last_img_fmt, &images, &capacity, &count, block_end, last_img_start);
            }
            buffer_collector_free(&mso_collector);
            collecting_mso = 0;

            last_img_start = NULL;
            last_img_fmt = XLS2IMG_UNKNOWN;
        }

        ptr += recordSize;
        if (ptr > end_ptr)
            break;
    }
    if (count > 0)
    {
        if (capacity > count * 2)
        {
            XLS2IMG_IMAGE* new_images = (XLS2IMG_IMAGE*)realloc(images, count * sizeof(XLS2IMG_IMAGE));
            if (new_images)
                images = new_images;
        }

        result->images = images;
        result->count = count;
        return count;
    }
    else
    {
        free(images);
        return XLS2IMG_ERROR_NO_IMAGES;
    }
}

void xls2img_free_result(XLS2IMG_RESULT* result)
{
    if (!result) return;

    if (result->images)
    {
        for (int i = 0; i < result->count; i++)
            if (result->images[i].data)
                free(result->images[i].data);

        free(result->images);
        result->images = NULL;
    }
    result->count = 0;
}