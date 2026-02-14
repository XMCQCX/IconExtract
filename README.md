# IconExtract
Extract icons from resource files to ICO or PNG.

---

## Requirements
- AutoHotkey v2
- [GDI+ library](https://github.com/buliasz/AHKv2-Gdip/blob/master/Gdip_All.ahk) (must be initialized externally)

--- 

## Class Methods

### `SaveIconToIco(filePath, groupIndex, outputPath)`
Save an icon group as a multi-resolution .ico file.

---  

### `SaveIconToPng(filePath, groupIndex, outputPath)`
Save the highest quality icon variant as a .png file.

---  

### `SaveAllIconsToIco(filePath, outputDir)`
Save all icon groups as multi-resolution .ico files.

---  

### `SaveAllIconsToPng(filePath, outputDir)`
Save the highest quality icon variant from each icon group as .png files.

---  

### `GetIconGroupIds(filePath)`
Get all icon group resource identifiers.

---  

### `GetIconVariants(filePath, resourceName, copyData)`
Get all icon variants in an icon group.

---  

### `GetIconVariantCount(filePath, groupIndex)`
Get the number of icon variants in an icon group.

---  

### `GetIconGroupCount(filePath)`
Get the number of icon groups.

--- 

The `IconExtract.ahk` file includes JSDoc documentation for all method parameters. For usage examples, see `Examples.ahk`.

--- 

## License

- MIT License

## Credits
- [AutoHotkey](https://www.autohotkey.com) - Steve Gray, Chris Mallett, portions of the AutoIt Team, and various others.
- [AHKv2-Gdip](https://github.com/buliasz/AHKv2-Gdip/blob/master/Gdip_All.ahk) - tic (original [Gdip.ahk](https://github.com/tariqporter/Gdip)), Rseding91, mmikeww, buliasz, and various others.
