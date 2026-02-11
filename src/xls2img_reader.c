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

#pragma pack(push)
#pragma pack(1)

typedef struct {
    uint8_t signature[8];
    uint8_t unused_clsid[16];
    uint16_t minorVersion;
    uint16_t majorVersion;
    uint16_t byteOrder;
    uint16_t sectorShift;
    uint16_t miniSectorShift;
    uint8_t reserved[6];
    uint32_t numDirectorySector;
    uint32_t numFATSector;
    uint32_t firstDirectorySectorLocation;
    uint32_t transactionSignatureNumber;
    uint32_t miniStreamCutoffSize;
    uint32_t firstMiniFATSectorLocation;
    uint32_t numMiniFATSector;
    uint32_t firstDIFATSectorLocation;
    uint32_t numDIFATSector;
    uint32_t headerDIFAT[109];
} COMPOUND_FILE_HDR;

typedef struct {
    uint16_t name[32];
    uint16_t nameLen;
    uint8_t type;
    uint8_t colorFlag;
    uint32_t leftSiblingID;
    uint32_t rightSiblingID;
    uint32_t childID;
    uint8_t clsid[16];
    uint32_t stateBits;
    uint64_t creationTime;
    uint64_t modifiedTime;
    uint32_t startSectorLocation;
    uint64_t size;
} COMPOUND_FILE_ENTRY;

#pragma pack(pop)

struct XLS2IMG_READER {
    const uint8_t* buffer;
    size_t bufferLen;
    const COMPOUND_FILE_HDR* hdr;
    size_t sectorSize;
    size_t minisectorSize;
    size_t miniStreamStartSector;
};

static uint32_t parse_uint32(const void* buffer);
static uint32_t xls2img_get_fat_sector_location(const XLS2IMG_READER* reader, size_t fatSectorNumber);
static uint32_t xls2img_get_next_sector(const XLS2IMG_READER* reader, size_t sector);
static const uint8_t* xls2img_sector_offset_to_address(const XLS2IMG_READER* reader, size_t sector, size_t offset);
static void xls2img_locate_final_sector(const XLS2IMG_READER* reader, size_t sector, size_t offset,
    size_t* finalSector, size_t* finalOffset);
static void xls2img_read_stream(const XLS2IMG_READER* reader, size_t sector, size_t offset, char* buffer, size_t len);
static uint32_t xls2img_get_next_mini_sector(const XLS2IMG_READER* reader, size_t miniSector);
static const uint8_t* xls2img_mini_sector_offset_to_address(const XLS2IMG_READER* reader, size_t sector, size_t offset);
static void xls2img_locate_final_mini_sector(const XLS2IMG_READER* reader, size_t sector, size_t offset,
    size_t* finalSector, size_t* finalOffset);
static void xls2img_read_mini_stream(const XLS2IMG_READER* reader, size_t sector, size_t offset, char* buffer, size_t len);
static const COMPOUND_FILE_ENTRY* xls2img_get_entry(const XLS2IMG_READER* reader, uint32_t entryID);
static void xls2img_read_file(const XLS2IMG_READER* reader, const COMPOUND_FILE_ENTRY* entry, char* buffer);
static int xls2img_string_compare(const uint16_t* str1, const uint16_t* str2, size_t len);

const char* xls2img_strerror(int error_code)
{
    switch (error_code) 
    {
        case XLS2IMG_SUCCESS:                   return "Success";
        case XLS2IMG_ERROR_WRONG_FORMAT:        return "Wrong file format";
        case XLS2IMG_ERROR_FILE_CORRUPTED:      return "File is corrupted";
        case XLS2IMG_ERROR_INVALID_ARGUMENT:    return "Invalid argument";
        case XLS2IMG_ERROR_NO_WORKBOOK:         return "No workbook found";
        case XLS2IMG_ERROR_NO_IMAGES:           return "No images found";
        default:                                return "Unknown error";
    }
}

int xls2img_open(XLS2IMG_READER** reader, const void* buffer, size_t len)
{
    if (!buffer || len == 0) return XLS2IMG_ERROR_INVALID_ARGUMENT;

    XLS2IMG_READER* r = (XLS2IMG_READER*)malloc(sizeof(XLS2IMG_READER));
    if (!r) return XLS2IMG_ERROR_INVALID_ARGUMENT;

    r->buffer = (const uint8_t*)buffer;
    r->bufferLen = len;
    r->hdr = (const COMPOUND_FILE_HDR*)buffer;

    if (r->bufferLen < sizeof(COMPOUND_FILE_HDR) ||
        memcmp(r->hdr->signature, "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1", 8) != 0)
    {
        free(r);
        return XLS2IMG_ERROR_WRONG_FORMAT;
    }

    r->sectorSize = r->hdr->majorVersion == 3 ? 512 : 4096;
    r->minisectorSize = 64;

    if (r->bufferLen < r->sectorSize * 3)
    {
        free(r);
        return XLS2IMG_ERROR_FILE_CORRUPTED;
    }

    const COMPOUND_FILE_ENTRY* root = xls2img_get_entry(r, 0);
    if (!root)
    {
        free(r);
        return XLS2IMG_ERROR_FILE_CORRUPTED;
    }

    r->miniStreamStartSector = root->startSectorLocation;
    *reader = r;
    return XLS2IMG_SUCCESS;
}

void xls2img_close(XLS2IMG_READER* reader)
{
    if (reader)
        free(reader);
}

int xls2img_get_workbook(XLS2IMG_READER* reader, void** data, size_t* size)
{
    const COMPOUND_FILE_ENTRY* root = xls2img_get_entry(reader, 0);
    if (!root) return XLS2IMG_ERROR_FILE_CORRUPTED;

    uint32_t childID = root->childID;
    while (childID != 0xFFFFFFFF)
    {
        const COMPOUND_FILE_ENTRY* entry = xls2img_get_entry(reader, childID);
        if (!entry) break;

        if (entry->type == 2)
        {
            // stream type
            static const uint16_t workbookStr[] = { 'W','o','r','k','b','o','o','k',0 };
            static const uint16_t WORKBOOKStr[] = { 'W','O','R','K','B','O','O','K',0 };

            size_t nameLen = entry->nameLen / 2;
            if (nameLen > 0)
            {
                if (xls2img_string_compare(entry->name, workbookStr, nameLen) ||
                    xls2img_string_compare(entry->name, WORKBOOKStr, nameLen))
                {

                    char* workbook_data = (char*)malloc((size_t)entry->size);
                    if (!workbook_data) return XLS2IMG_ERROR_INVALID_ARGUMENT;

                    xls2img_read_file(reader, entry, workbook_data);
                    *data = workbook_data;
                    *size = (size_t)entry->size;
                    return XLS2IMG_SUCCESS;
                }
            }
        }

        if (entry->leftSiblingID != 0xFFFFFFFF)
            childID = entry->leftSiblingID;
        else if (entry->rightSiblingID != 0xFFFFFFFF)
            childID = entry->rightSiblingID;
        else
            break;
    }

    return XLS2IMG_ERROR_NO_WORKBOOK;
}

void xls2img_free_workbook_data(void* data)
{
    if (data)
        free(data);
}

static uint32_t parse_uint32(const void* buffer)
{
    return *((const uint32_t*)buffer);
}

static uint32_t xls2img_get_fat_sector_location(const XLS2IMG_READER* reader, size_t fatSectorNumber)
{
    if (fatSectorNumber < 109)
        return reader->hdr->headerDIFAT[fatSectorNumber];

    fatSectorNumber -= 109;
    size_t entriesPerSector = reader->sectorSize / 4 - 1;
    uint32_t difatSectorLocation = reader->hdr->firstDIFATSectorLocation;

    while (fatSectorNumber >= entriesPerSector)
    {
        fatSectorNumber -= entriesPerSector;
        const uint8_t* addr = xls2img_sector_offset_to_address(reader, difatSectorLocation, reader->sectorSize - 4);
        if (!addr) return 0xFFFFFFFF;
        difatSectorLocation = parse_uint32(addr);
    }

    const uint8_t* addr = xls2img_sector_offset_to_address(reader, difatSectorLocation, fatSectorNumber * 4);
    if (!addr) return 0xFFFFFFFF;

    return parse_uint32(addr);
}

static uint32_t xls2img_get_next_sector(const XLS2IMG_READER* reader, size_t sector)
{
    size_t entriesPerSector = reader->sectorSize / 4;
    size_t fatSectorNumber = sector / entriesPerSector;
    uint32_t fatSectorLocation = xls2img_get_fat_sector_location(reader, fatSectorNumber);

    const uint8_t* addr = xls2img_sector_offset_to_address(reader, fatSectorLocation, (sector % entriesPerSector) * 4);
    if (!addr) return 0xFFFFFFFF;

    return parse_uint32(addr);
}

static const uint8_t* xls2img_sector_offset_to_address(const XLS2IMG_READER* reader, size_t sector, size_t offset)
{
    if (sector >= 0xFFFFFFFA || offset >= reader->sectorSize) return NULL;

    size_t pos = reader->sectorSize + reader->sectorSize * sector + offset;
    if (pos >= reader->bufferLen) return NULL;

    return reader->buffer + pos;
}

static void xls2img_locate_final_sector(const XLS2IMG_READER* reader, size_t sector, size_t offset, size_t* finalSector, size_t* finalOffset)
{
    while (offset >= reader->sectorSize)
    {
        offset -= reader->sectorSize;
        uint32_t next = xls2img_get_next_sector(reader, sector);
        if (next == 0xFFFFFFFE || next == 0xFFFFFFFF) break;
        sector = next;
    }
    *finalSector = sector;
    *finalOffset = offset;
}

static void xls2img_read_stream(const XLS2IMG_READER* reader, size_t sector, size_t offset, char* buffer, size_t len) {
    xls2img_locate_final_sector(reader, sector, offset, &sector, &offset);

    while (len > 0)
    {
        const uint8_t* src = xls2img_sector_offset_to_address(reader, sector, offset);
        if (!src) break;

        size_t copylen = len < reader->sectorSize - offset ? len : reader->sectorSize - offset;
        if (reader->buffer + reader->bufferLen < src + copylen) break;

        memcpy(buffer, src, copylen);
        buffer += copylen;
        len -= copylen;
        sector = xls2img_get_next_sector(reader, sector);
        offset = 0;

        if (sector == 0xFFFFFFFE || sector == 0xFFFFFFFF) break;
    }
}

static uint32_t xls2img_get_next_mini_sector(const XLS2IMG_READER* reader, size_t miniSector)
{
    size_t sector, offset;
    xls2img_locate_final_sector(reader, reader->hdr->firstMiniFATSectorLocation, miniSector * 4, &sector, &offset);
    const uint8_t* addr = xls2img_sector_offset_to_address(reader, sector, offset);
    if (!addr) return 0xFFFFFFFF;
    return parse_uint32(addr);
}

static const uint8_t* xls2img_mini_sector_offset_to_address(const XLS2IMG_READER* reader, size_t sector, size_t offset)
{
    size_t sectorOffset, sectorPos;
    xls2img_locate_final_sector(reader, reader->miniStreamStartSector,
        sector * reader->minisectorSize + offset, &sectorPos, &sectorOffset);
    return xls2img_sector_offset_to_address(reader, sectorPos, sectorOffset);
}

static void xls2img_locate_final_mini_sector(const XLS2IMG_READER* reader, size_t sector, size_t offset, size_t* finalSector, size_t* finalOffset)
{
    while (offset >= reader->minisectorSize)
    {
        offset -= reader->minisectorSize;
        uint32_t next = xls2img_get_next_mini_sector(reader, sector);
        if (next == 0xFFFFFFFE || next == 0xFFFFFFFF) break;
        sector = next;
    }
    *finalSector = sector;
    *finalOffset = offset;
}

static void xls2img_read_mini_stream(const XLS2IMG_READER* reader, size_t sector, size_t offset, char* buffer, size_t len)
{
    xls2img_locate_final_mini_sector(reader, sector, offset, &sector, &offset);

    while (len > 0)
    {
        const uint8_t* src = xls2img_mini_sector_offset_to_address(reader, sector, offset);
        if (!src) break;

        size_t copylen = len < reader->minisectorSize - offset ? len : reader->minisectorSize - offset;
        if (reader->buffer + reader->bufferLen < src + copylen) break;

        memcpy(buffer, src, copylen);
        buffer += copylen;
        len -= copylen;
        sector = xls2img_get_next_mini_sector(reader, sector);
        offset = 0;

        if (sector == 0xFFFFFFFE || sector == 0xFFFFFFFF) break;
    }
}

static const COMPOUND_FILE_ENTRY* xls2img_get_entry(const XLS2IMG_READER* reader, uint32_t entryID)
{
    if (entryID == 0xFFFFFFFF) return NULL;

    size_t sector = 0;
    size_t offset = 0;
    xls2img_locate_final_sector(reader, reader->hdr->firstDirectorySectorLocation,
        entryID * sizeof(COMPOUND_FILE_ENTRY), &sector, &offset);

    if (reader->bufferLen <= (size_t)(reader->sectorSize + reader->sectorSize * sector + offset + sizeof(COMPOUND_FILE_ENTRY)))
        return NULL;

    return (const COMPOUND_FILE_ENTRY*)(reader->buffer + reader->sectorSize + reader->sectorSize * sector + offset);
}

static void xls2img_read_file(const XLS2IMG_READER* reader, const COMPOUND_FILE_ENTRY* entry, char* buffer)
{
    if (entry->size == 0) return;

    if (entry->size < reader->hdr->miniStreamCutoffSize)
        xls2img_read_mini_stream(reader, entry->startSectorLocation, 0, buffer, (size_t)entry->size);
    else
        xls2img_read_stream(reader, entry->startSectorLocation, 0, buffer, (size_t)entry->size);
}

static int xls2img_string_compare(const uint16_t* str1, const uint16_t* str2, size_t len)
{
    for (size_t i = 0; i < len; i++) {
        if (str1[i] != str2[i]) return 0;
        if (str1[i] == 0) break;
    }
    return 1;
}