/*********************************************************************************************
 * IconExtract - Extract icons from resource files to ICO or PNG
 *
 * Requires Gdip_All.ahk library and GDI+ to be initialized externally.
 * {@link https://github.com/buliasz/AHKv2-Gdip/blob/master/Gdip_All.ahk Gdip_All.ahk}
 *
 * @author Martin Chartier (XMCQCX)
 * @version 1.0.0
 * @date 2026/02/13
 * @license MIT
 * @credits
 * - icons_ahk2 by TheArkive. {@link https://github.com/TheArkive/icons_ahk2 GitHub}
 *
 * Available Methods:
 * - SaveIconToIco       : Save an icon group as a multi-resolution .ico file.
 * - SaveIconToPng       : Save the highest quality icon variant as a .png file.
 * - SaveAllIconsToIco   : Save all icon groups as multi-resolution .ico files.
 * - SaveAllIconsToPng   : Save the highest quality icon variant from each icon group as .png files.
 * - GetIconGroupIds     : Get all icon group resource identifiers.
 * - GetIconVariants     : Get all icon variants in an icon group.
 * - GetIconVariantCount : Get the number of icon variants in an icon group.
 * - GetIconGroupCount   : Get the number of icon groups.
 */
Class IconExtract {
    ; Resource type constants
    static RT_ICON := 3
    static RT_GROUP_ICON := 14

    ; PNG format constants
    static PNG_SIGNATURE_SIZE := 4
    static PNG_IHDR_MIN_SIZE := 24
    static PNG_IHDR_WIDTH_OFFSET := 16
    static PNG_IHDR_HEIGHT_OFFSET := 20

    ; ICO format constants
    static ICO_HEADER_SIZE := 6
    static ICO_ENTRY_SIZE := 16
    static ICON_GROUP_ENTRY_SIZE := 14
    static ICO_TYPE_OFFSET := 2
    static ICO_TYPE_ICON := 1

    ; Icon dimension constants
    static ICON_DIM_NORMALIZED := 256

    /*********************************************************************************************
     * Save an icon group as a multi-resolution .ico file.
     *
     * @param {string} filePath        - Path to the resource file.
     * @param {integer} [groupIndex=1] - Icon group index.
     * @param {string} [outputPath=""] - Optional output path. Can be:
     *                                   - Omitted: auto-generates in A_ScriptDir
     *                                   - Directory path: auto-generates filename in that directory
     *                                   - Full file path: uses as-is
     *
     * @returns {string}               - Full path of the saved .ico file.
     ********************************************************************************************/
    static SaveIconToIco(filePath, groupIndex := 1, outputPath := '')
    {
        groupData := this._getIconsFromGroup(filePath, groupIndex, true)
        return this._saveIconsToIcoFile(filePath, groupData.icons, groupIndex, outputPath)
    }

    /*********************************************************************************************
     * Save the highest quality icon variant as a .png file.
     *
     * @param {string} filePath        - Path to the resource file.
     * @param {integer} [groupIndex=1] - Icon group index.
     * @param {string} [outputPath=""] - Optional output path. Can be:
     *                                   - Omitted: auto-generates in A_ScriptDir
     *                                   - Directory path: auto-generates filename in that directory
     *                                   - Full file path: uses as-is
     *
     * @returns {string}               - Full path of the saved .png file.
     ********************************************************************************************/
    static SaveIconToPng(filePath, groupIndex := 1, outputPath := '')
    {
        groupData := this._getIconsFromGroup(filePath, groupIndex, true)
        return this._saveLargestIconToPngFile(filePath, groupData.icons, groupData.resourceName, groupIndex, outputPath)
    }

    /*********************************************************************************************
     * Save all icon groups as multi-resolution .ico files.
     *
     * @param {string} filePath       - Path to the resource file.
     * @param {string} [outputDir=""] - Directory to save icons (default: `A_ScriptDir\<filename>-<ext>-ico`).
     *
     * @returns {array}               - Array of full paths to successfully saved .ico files.
     ********************************************************************************************/
    static SaveAllIconsToIco(filePath, outputDir := '')
    {
        return this._saveAllIcons(filePath, outputDir, 'ico')
    }

    /*********************************************************************************************
     * Save the highest quality icon variant from each icon group as .png files.
     *
     * @param {string} filePath       - Path to the resource file.
     * @param {string} [outputDir=""] - Directory to save PNG files (default: `A_ScriptDir\<filename>-<ext>-png`).
     *
     * @returns {array}               - Array of full paths to successfully saved .png files.
     ********************************************************************************************/
    static SaveAllIconsToPng(filePath, outputDir := '')
    {
        return this._saveAllIcons(filePath, outputDir, 'png')
    }

    /*********************************************************************************************
     * Extract and save all icon groups from a resource file into individual files.
     *
     * @param {string} filePath  - Path to the resource file.
     * @param {string} outputDir - Directory to save icons (empty for auto-generation).
     * @param {string} extension - File extension ('ico' or 'png').
     *
     * @returns {array}          - Array of full paths to successfully saved files.
     ********************************************************************************************/
    static _saveAllIcons(filePath, outputDir, extension)
    {
        iconGroups := this.GetIconGroupIds(filePath)

        if (outputDir = '') {
            SplitPath(filePath,,, &srcExt, &name)
            outputDir := A_ScriptDir '\' name '-' srcExt '-' extension
        }

        extractedFiles := []
        lastError := ''

        for (groupIndex, resourceName in iconGroups) {
            try {
                if (outputPath := this._saveIconByResource(filePath, resourceName, groupIndex, outputDir, extension))
                    extractedFiles.Push(outputPath)
            } catch as err {
                lastError := err
                continue
            }
        }

        if (!extractedFiles.Length) {
            errorMsg := 'All icon extractions failed.`nPath: "' filePath '"`nTotal groups: ' iconGroups.Length
            if (lastError)
                errorMsg .= '`n`nLast error: ' lastError.Message
            throw Error(errorMsg, -1)
        }

        return extractedFiles
    }

    /*********************************************************************************************
     * Get all icon group resource identifiers.
     *
     * @param {string} filePath - Path to the resource file.
     *
     * @returns {array}         - Array of group resource identifiers (string name or integer ID).
     ********************************************************************************************/
    static GetIconGroupIds(filePath)
    {
        if (!hModule := DllCall('LoadLibraryEx', 'Str', filePath, 'UPtr', 0, 'UInt', 0x2, 'UPtr'))
            throw Error('LoadLibraryEx failed. File may not exist, be inaccessible, or not contain resources.`nPath: "' filePath '"', -1)

        try {
            iconGroups := []

            EnumCallback := (hMod, lpType, lpName, lParam) => (
                iconGroups.Push((lpName >> 16 = 0) ? lpName : StrGet(lpName)),
                true
            )

            cb := CallbackCreate(EnumCallback, 'F', 4)
            try DllCall('EnumResourceNames', 'UPtr', hModule, 'Ptr', this.RT_GROUP_ICON, 'UPtr', cb, 'UPtr', 0, 'Int')
            finally
                CallbackFree(cb)

            if (!iconGroups.Length)
                throw ValueError('No icon groups found in file.`nPath: "' filePath '"', -1)

            return iconGroups
        }
        finally
            DllCall('FreeLibrary', 'UPtr', hModule)
    }

    /*********************************************************************************************
     * Get all icon variants in an icon group.
     *
     * @param {string} filePath             - Path to the resource file.
     * @param {string|integer} resourceName - Icon group resource identifier (string name or integer ID).
     * @param {boolean} [copyData=false]    - Whether to include raw icon data buffers.
     *
     * @returns {array}                     - Array of icon objects with properties:
     *                                        {type, width, height, colorCount, planes, bits, size, nID, data}
     ********************************************************************************************/
    static GetIconVariants(filePath, resourceName, copyData := false)
    {
        if !hModule := DllCall('LoadLibraryEx', 'Str', filePath, 'UPtr', 0, 'UInt', 0x2, 'UPtr')
            throw Error('LoadLibraryEx failed. File may not exist, be inaccessible, or not contain resources.`nPath: "' filePath '"', -1)

        try {
            groupInfoBuffer := this._getResourceBuffer(hModule, this.RT_GROUP_ICON, filePath, resourceName)

            if (!groupInfoBuffer || NumGet(groupInfoBuffer, this.ICO_TYPE_OFFSET, 'UShort') != this.ICO_TYPE_ICON)
                throw ValueError('Invalid or missing icon group data.`nPath: "' filePath '"`nResource Identifier: ' resourceName, -1)

            iconCount := NumGet(groupInfoBuffer, 4, 'UShort')

            if (iconCount = 0)
                throw ValueError('Icon group contains no icons.`nPath: "' filePath '"`nResource Identifier: ' resourceName, -1)

            offset := this.ICO_HEADER_SIZE
            icons := []

            Loop iconCount {
                icon := {type: 'dib'}
                icon.width := NumGet(groupInfoBuffer, offset, 'UChar')
                icon.height := NumGet(groupInfoBuffer, offset + 1, 'UChar')
                icon.colorCount := NumGet(groupInfoBuffer, offset + 2, 'UChar')
                icon.planes := NumGet(groupInfoBuffer, offset + 4, 'UShort')
                icon.bits := NumGet(groupInfoBuffer, offset + 6, 'UShort')
                icon.size := NumGet(groupInfoBuffer, offset + 8, 'UInt')
                icon.nID := NumGet(groupInfoBuffer, offset + 12, 'UShort')

                iconDataBuffer := this._getResourceBuffer(hModule, this.RT_ICON, filePath, icon.nID)

                ; PNG-compressed: read actual dimensions from IHDR chunk
                if (iconDataBuffer && iconDataBuffer.Size >= this.PNG_IHDR_MIN_SIZE && this._isPngCompressed(iconDataBuffer)) {
                    icon.width := this._readBigEndianUInt32(iconDataBuffer, this.PNG_IHDR_WIDTH_OFFSET)
                    icon.height := this._readBigEndianUInt32(iconDataBuffer, this.PNG_IHDR_HEIGHT_OFFSET)
                    icon.type := 'png'
                }
                ; BMP-based (DIB): keep raw header values (0=256), normalize with _normalizeIconDimensions when required

                icon.data := (copyData ? iconDataBuffer : '')
                offset += this.ICON_GROUP_ENTRY_SIZE
                icons.Push(icon)
            }

            return icons
        }
        finally
            DllCall('FreeLibrary', 'UPtr', hModule)
    }

    /*********************************************************************************************
     * Get the number of icon variants in an icon group.
     *
     * @param {string} filePath    - Path to the resource file.
     * @param {integer} groupIndex - Icon group index.
     *
     * @returns {integer}          - Number of variants (different sizes/color depths) in the group.
     ********************************************************************************************/
    static GetIconVariantCount(filePath, groupIndex)
    {
        groupData := this._getIconsFromGroup(filePath, groupIndex, false)
        return groupData.icons.Length
    }

    /*********************************************************************************************
     * Get the number of icon groups.
     *
     * @param {string} filePath - Path to the resource file.
     *
     * @returns {integer}       - Number of icon groups in the file.
     ********************************************************************************************/
    static GetIconGroupCount(filePath)
    {
        return this.GetIconGroupIds(filePath).Length
    }

    /*********************************************************************************************
     * Get icons from a group.
     *
     * @param {string} filePath    - Path to the resource file.
     * @param {integer} groupIndex - Icon group index.
     * @param {boolean} copyData   - Whether to include icon data buffers.
     *
     * @returns {object}           - Object containing {icons: array, resourceName: string|integer, groupIndex: integer}
     ********************************************************************************************/
    static _getIconsFromGroup(filePath, groupIndex, copyData)
    {
        iconGroups := this.GetIconGroupIds(filePath)

        if (!IsInteger(groupIndex) || groupIndex < 1)
            throw TypeError('Invalid group index. Expected an integer greater than or equal to 1.', -1)

        if (groupIndex > iconGroups.Length)
            throw ValueError('Invalid group index.`nPath: "' filePath '"`nGroup index: ' groupIndex '`nAvailable groups: ' iconGroups.Length, -1)

        resourceName := iconGroups[groupIndex]
        icons := this.GetIconVariants(filePath, resourceName, copyData)

        return {
            icons: icons,
            resourceName: resourceName,
            groupIndex: groupIndex
        }
    }

    /*********************************************************************************************
     * Save an icon using resourceName directly.
     *
     * @param {string} filePath             - Path to the resource file.
     * @param {string|integer} resourceName - Icon group resource identifier (string name or integer ID).
     * @param {integer} groupIndex          - Icon group index.
     * @param {string} outputPath           - Output path or directory.
     * @param {string} extension            - File extension ('ico' or 'png').
     *
     * @returns {string}                    - Full path of the saved file.
     ********************************************************************************************/
    static _saveIconByResource(filePath, resourceName, groupIndex, outputPath, extension)
    {
        icons := this.GetIconVariants(filePath, resourceName, true)

        if (extension == 'ico')
            return this._saveIconsToIcoFile(filePath, icons, groupIndex, outputPath)
        else
            return this._saveLargestIconToPngFile(filePath, icons, resourceName, groupIndex, outputPath)
    }

    /*********************************************************************************************
     * Core ICO file writing logic shared by SaveIconToIco and _saveIconByResource.
     *
     * @param {string} filePath    - Path to the resource file.
     * @param {array} icons        - Array of icon objects with data.
     * @param {integer} groupIndex - Icon group index.
     * @param {string} outputPath  - Output path or directory.
     *
     * @returns {string}           - Full path of the saved .ico file.
     ********************************************************************************************/
    static _saveIconsToIcoFile(filePath, icons, groupIndex, outputPath)
    {
        outputPath := this._resolveOutputPath(filePath, groupIndex, icons, 'ico', outputPath)
        SplitPath(outputPath,,, &ext)

        if (ext != 'ico')
            throw Error('Invalid file extension ".' ext '". Must be ".ico".`nPath: "' outputPath '"', -1)

        try file := FileOpen(outputPath, 'w')
        catch as err
            throw Error('Cannot write to output file. ' err.Message '`nPath: "' outputPath '"' , -1)

        file.RawWrite(this._buildIcoHeader(icons))
        Loop icons.Length
            file.RawWrite(icons[A_Index].data)

        file.Close()
        return outputPath
    }

    /*********************************************************************************************
     * Determines icon format (PNG or BMP) and routes to appropriate save method.
     *
     * @param {string} filePath             - Path to the resource file.
     * @param {array} icons                 - Array of icon objects with data.
     * @param {string|integer} resourceName - Icon group resource identifier (string name or integer ID).
     * @param {integer} groupIndex          - Icon group index.
     * @param {string} outputPath           - Output path or directory.
     *
     * @returns {string}                    - Full path of the saved .png file.
     ********************************************************************************************/
    static _saveLargestIconToPngFile(filePath, icons, resourceName, groupIndex, outputPath)
    {
        largestIcon := this._getLargestIcon(icons)
        largestIconBuffer := largestIcon.data
        outputPath := this._resolveOutputPath(filePath, groupIndex, largestIcon, 'png', outputPath)
        SplitPath(outputPath,,, &ext)

        if (ext != 'png')
            throw Error('Invalid file extension ".' ext '". Must be ".png".`nPath: "' outputPath '"', -1)

        if this._isPngCompressed(largestIconBuffer)
            return this._savePngIconToPng(largestIconBuffer, outputPath)

        return this._saveBmpIconToPng(filePath, resourceName, groupIndex, outputPath, largestIcon)
    }

    /*********************************************************************************************
     * Checks if icon data uses PNG compression by verifying PNG signature.
     *
     * @param {Buffer} iconDataBuffer - Buffer containing icon data.
     *
     * @returns {boolean}             - True if data starts with PNG signature.
     ********************************************************************************************/
    static _isPngCompressed(iconDataBuffer)
    {
        if (!iconDataBuffer || iconDataBuffer.Size < this.PNG_SIGNATURE_SIZE)
            return false

        ; Check for PNG signature: 0x89 0x50 0x4E 0x47
        return (NumGet(iconDataBuffer, 0, 'UChar') = 0x89
             && NumGet(iconDataBuffer, 1, 'UChar') = 0x50
             && NumGet(iconDataBuffer, 2, 'UChar') = 0x4E
             && NumGet(iconDataBuffer, 3, 'UChar') = 0x47)
    }

    /*********************************************************************************************
     * Reads a 32-bit unsigned integer in big-endian format from a buffer and returns the PNG IHDR chunk width or height.
     *
     * @param {Buffer} buffer  - Source buffer.
     * @param {integer} offset - Byte offset to read from.
     *
     * @returns {integer}      - The 32-bit value.
     ********************************************************************************************/
    static _readBigEndianUInt32(buffer, offset)
    {
        if (!buffer || buffer.Size < offset + 4)
            throw ValueError('Buffer too small for reading 32-bit value at offset ' offset, -1)

        return (NumGet(buffer, offset, 'UChar') << 24)
             | (NumGet(buffer, offset + 1, 'UChar') << 16)
             | (NumGet(buffer, offset + 2, 'UChar') << 8)
             | NumGet(buffer, offset + 3, 'UChar')
    }

    /*********************************************************************************************
     * Gets the icon with the largest dimensions and highest color depth.
     *
     * @param {array} icons - Array of icon objects.
     *
     * @returns {object}    - Icon object with the largest area and best quality.
     ********************************************************************************************/
    static _getLargestIcon(icons)
    {
        if (!icons.Length)
            throw ValueError('No icons provided in array', -1)

        maxArea := 0
        maxBits := 0
        largestIcon := ''

        for (icon in icons) {
            dimensions := this._normalizeIconDimensions(icon)
            area := dimensions.w * dimensions.h

            if (area > maxArea || (area = maxArea && icon.bits > maxBits)) {
                maxArea := area
                maxBits := icon.bits
                largestIcon := icon
            }
        }

        return largestIcon
    }

    /*********************************************************************************************
     * Normalizes icon dimensions, treating 0 as 256.
     *
     * @param {object} icon - Icon object with width and height properties (raw values from header).
     *
     * @returns {object}    - Object with normalized {w, h} properties (actual pixel dimensions).
     ********************************************************************************************/
    static _normalizeIconDimensions(icon)
    {
        return {
            w: (icon.width = 0) ? this.ICON_DIM_NORMALIZED : icon.width,
            h: (icon.height = 0) ? this.ICON_DIM_NORMALIZED : icon.height
        }
    }

    /*********************************************************************************************
     * Gets dimensions from either a single icon or the highest quality variant in an array.
     *
     * @param {object|array} iconOrIcons - Single icon object or array of icons.
     *
     * @returns {object}                 - Object with {w, h} properties (normalized dimensions).
     ********************************************************************************************/
    static _getDimensions(iconOrIcons)
    {
        icon := (iconOrIcons is Array) ? this._getLargestIcon(iconOrIcons) : iconOrIcons
        return this._normalizeIconDimensions(icon)
    }

    /*********************************************************************************************
     * Generates filename for icon files (without directory path).
     *
     * @param {string} filePath          - Path to the resource file.
     * @param {integer} groupIndex       - Icon group index.
     * @param {object|array} iconOrIcons - Single icon object (for PNG) or array of icons (for ICO).
     * @param {string} extension         - 'ico' or 'png'.
     *
     * @returns {string}                 - Filename in format: `<filename>-<ext>-<groupIndex>-<size>.<extension>`
     *                                     Example: "shell32-dll-1-256x256.png"
     ********************************************************************************************/
    static _generateFileName(filePath, groupIndex, iconOrIcons, extension)
    {
        SplitPath(filePath,,, &srcExt, &fileNameNoExt)
        dimensions := this._getDimensions(iconOrIcons)
        sizeStr := dimensions.w 'x' dimensions.h
        return fileNameNoExt '-' srcExt '-' groupIndex '-' sizeStr '.' extension
    }

    /*********************************************************************************************
     * Resolves output path for icon files, handling empty strings, directories, and file paths.
     *
     * @param {string} filePath          - Path to the resource file.
     * @param {integer} groupIndex       - Icon group index.
     * @param {object|array} iconOrIcons - Single icon object (for PNG) or array of icons (for ICO).
     * @param {string} extension         - 'ico' or 'png'.
     * @param {string} outputPath        - User-provided output path (empty, directory, or file).
     *
     * @returns {string}                 - Full resolved path to output file.
     *********************************************************************************************/
    static _resolveOutputPath(filePath, groupIndex, iconOrIcons, extension, outputPath)
    {
        fileName := this._generateFileName(filePath, groupIndex, iconOrIcons, extension)

        if (outputPath = '')
            resolvedPath := A_ScriptDir '\' fileName
        else if (DirExist(outputPath))
            resolvedPath := outputPath '\' fileName
        else {
            SplitPath(outputPath,,, &ext)
            resolvedPath := (ext = '') ? outputPath '\' fileName : outputPath
        }

        ; Create parent directory if it doesn't exist
        SplitPath(resolvedPath,, &dir)

        if (dir && !DirExist(dir)) {
            try DirCreate(dir)
            catch as err
                throw Error('Cannot create output directory. ' err.Message '`nPath: "' dir '"' , -1)
        }

        return resolvedPath
    }

    /*********************************************************************************************
     * Saves PNG-compressed icon data directly to file (no conversion needed).
     *
     * @param {Buffer} iconDataBuffer - Buffer containing raw PNG data.
     * @param {string} outputPath     - Full path for the output .png file.
     *
     * @returns {string}              - Path of the saved file.
     ********************************************************************************************/
    static _savePngIconToPng(iconDataBuffer, outputPath)
    {
        try file := FileOpen(outputPath, 'w')
        catch as err
            throw Error('Cannot write to output file. ' err.Message '`nPath: "' outputPath '"' , -1)

        file.RawWrite(iconDataBuffer)
        file.Close()
        return outputPath
    }

    /*********************************************************************************************
     * Converts a BMP-based (DIB format) icon to PNG using CreateIconFromResourceEx and GDI+.
     *
     * @param {string} filePath             - Path to the resource file.
     * @param {string|integer} resourceName - Icon group resource identifier (string name or integer ID).
     * @param {integer} groupIndex          - Icon group index.
     * @param {string} outputPath           - Full path for the output .png file.
     * @param {object} largestIcon          - Icon object containing metadata and raw data buffer.
     *
     * @returns {string}                    - Path of the saved .png file.
     ********************************************************************************************/
    static _saveBmpIconToPng(filePath, resourceName, groupIndex, outputPath, largestIcon)
    {
        hIcon := DllCall('CreateIconFromResourceEx',
            'Ptr', largestIcon.data.Ptr,
            'UInt', largestIcon.data.Size,
            'Int', 1,
            'UInt', 0x00030000,
            'Int', 0,
            'Int', 0,
            'UInt', 0,
            'Ptr')

        if (!hIcon)
            throw Error('CreateIconFromResourceEx failed.`nPath: "' filePath '"`nGroup index: ' groupIndex '`nResource Identifier: ' resourceName, -1)

        pBitmap := Gdip_CreateBitmapFromHICON(hIcon)
        DllCall('DestroyIcon', 'Ptr', hIcon)

        if (!pBitmap) {
            if (largestIcon.bits = 1)
                throw Error('Cannot convert monochrome 1-bit icon to PNG (GDI+ limitation).`nPath: "' filePath '"`nGroup index: ' groupIndex '`nResource Identifier: ' resourceName, -1)

            throw Error('Gdip_CreateBitmapFromHICON failed.`nPath: "' filePath '"`nGroup index: ' groupIndex '`nResource Identifier: ' resourceName
                    '`n`nPossible causes:`n- GDI+ not initialized (call Gdip_Startup() first)`n- Corrupted or invalid icon data', -1)
        }

        result := Gdip_SaveBitmapToFile(pBitmap, outputPath)
        Gdip_DisposeImage(pBitmap)

        if (result != 0) {
            switch result {
                case -1: errMsg := 'Extension supplied is not a supported file format.'
                case -2: errMsg := 'Could not get a list of encoders on system.'
                case -3: errMsg := 'Could not find matching encoder for specified file format.'
                case -4: errMsg := 'Could not get WideChar name of output file.'
                case -5: errMsg := 'Could not save file to disk.'
                default: errMsg := 'Unknown error code: ' result '.'
            }
            throw Error('Gdip_SaveBitmapToFile failed. ' errMsg '`nPath: "' outputPath '"', -1)
        }

        return outputPath
    }

    /*********************************************************************************************
     * Builds the standard .ico file header structure.
     *
     * @param {array} icons - Array of icon objects (with raw dimension values from headers).
     *
     * @returns {Buffer}    - Buffer containing the .ico file header (ICONDIR structure).
     ********************************************************************************************/
    static _buildIcoHeader(icons)
    {
        count := icons.Length
        headerSize := this.ICO_HEADER_SIZE + this.ICO_ENTRY_SIZE * count
        iconHeaderBuffer := Buffer(headerSize, 0)

        NumPut('UShort', 0, iconHeaderBuffer, 0)      ; reserved
        NumPut('UShort', 1, iconHeaderBuffer, 2)      ; type = ICO
        NumPut('UShort', count, iconHeaderBuffer, 4)  ; icon count

        offset := this.ICO_HEADER_SIZE
        dataOffset := headerSize

        Loop count {
            icon := icons[A_Index]

            NumPut('UChar', icon.width, iconHeaderBuffer, offset)
            NumPut('UChar', icon.height, iconHeaderBuffer, offset + 1)
            NumPut('UChar', icon.colorCount, iconHeaderBuffer, offset + 2)
            NumPut('UChar', 0, iconHeaderBuffer, offset + 3)  ; reserved
            NumPut('UShort', icon.planes, iconHeaderBuffer, offset + 4)
            NumPut('UShort', icon.bits, iconHeaderBuffer, offset + 6)
            NumPut('UInt', icon.size, iconHeaderBuffer, offset + 8)
            NumPut('UInt', dataOffset, iconHeaderBuffer, offset + 12)

            offset += this.ICO_ENTRY_SIZE
            dataOffset += icon.size
        }

        return iconHeaderBuffer
    }

    /*********************************************************************************************
     * Gets a resource buffer from a loaded module.
     *
     * @param {UPtr} hModule                - Handle to the loaded module (from LoadLibraryEx).
     * @param {integer} resourceType        - Resource type (RT_ICON or RT_GROUP_ICON).
     * @param {string} filePath             - Path to the resource file.
     * @param {string|integer} resourceName - Icon group resource identifier (string name or integer ID).
     *
     * @returns {Buffer}                    - Buffer containing the resource data.
     ********************************************************************************************/
    static _getResourceBuffer(hModule, resourceType, filePath, resourceName)
    {
        resourceIdentifier := (IsInteger(resourceName) ? '#' resourceName : resourceName)

        if (!hRes := DllCall('FindResource', 'UPtr', hModule, 'Str', resourceIdentifier, 'Ptr', resourceType, 'UPtr'))
            throw ValueError('FindResource failed. Resource not found or inaccessible.`nPath: "' filePath '"`nResource Identifier: ' resourceName '`nType: ' resourceType, -1)

        if (!hLoaded := DllCall('LoadResource', 'UPtr', hModule, 'UPtr', hRes, 'UPtr'))
            throw Error('LoadResource failed. Could not load the resource data.', -1)

        if (!pData := DllCall('LockResource', 'UPtr', hLoaded, 'UPtr'))
            throw Error('LockResource failed. Could not lock the resource memory.', -1)

        size := DllCall('SizeofResource', 'UPtr', hModule, 'UPtr', hRes, 'UInt')
        if (size = 0)
            throw Error('SizeofResource returned 0. Resource is empty or corrupted.', -1)

        resourceDataBuffer := Buffer(size)
        DllCall('RtlCopyMemory', 'UPtr', resourceDataBuffer.Ptr, 'UPtr', pData, 'UPtr', size)
        return resourceDataBuffer
    }
}