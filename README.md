# LibreOffice Calc template for Anki

This repository contains LibreOffice Calc file (.odt) which contains convenience macros for quickly filling the required fields for Anki cards used for learning Japanese.

## Table of Contents
- [Available functions](#available-functions)
    - [jisho module](#jisho-module)
    - [ruby module](#ruby-module)
- [How to use the template](#how-to-use-the-template)
    - [Default columns & configuration](#default-columns--configuration)
        - [Default constant values](#default-constant-values)
        - [Columns](#columns)
    - [Demo](#demo)
    - [Initial Setup](#initial-setup)
- [Importing to Anki](#importing-to-anki)
- [Useful resources](#useful-resources)

## Available functions

### jisho module
* `CREATEDROPDOWNONCHANGE(oEvent As Variant)`

    Is used for "Content Changed" Sheet Event binding. 
    
    It checks whether the event was triggered in one of the cells of the configured column and if so, sends the value of the cell to Jisho. Once the response is available, it creates a dropdown with values parsed from Jisho results.

    The behavior can be configured in jisho macros module by changing the values of these constants:
    * ROMAJI_COLUMN - set to the number of the column in which cells should be tracked for change and their values sent to Jisho;
    * JISHO_RESULT_LIMIT - set to the number of results that should be returned from Jisho. The generated dropdown will have this many rows if there are enough results.


* `KANJI(TEXT as String)`

    Is used for parsing kanji from the value selected in the dropdown. The parameter should be set to the reference of the dropdown cell.

* `READING(TEXT as String)`

    Is used for parsing hiragana/katakana from the value selected in the dropdown. The parameter should be set to the reference of the dropdown cell.

### ruby module

* `RUBY(TEXT as String, ANNOTATION as String)`

    Used for creating html \<ruby\> tag (furigana).

* `RB(TEXT as String, SEPARATOR as String)`

    Used for inserting html \<ruby\> tags (furigana) inside a longer text, especially if one than more ruby element is needed. It finds patterns like this: `<SEPARATOR>kanji<SEPARATOR>hiraganaOrKatakana<SEPARATOR>` and replaces them with ruby elements. E.g. TEXT could be:
    ```
     |答|こた|えが|分|わ|かる
     ```
     With the SEPARATOR `|` the output is:
     ```
     <ruby>答<rt>こた</rt></ruby>えが<ruby>分<rt>わ</rt></ruby>かる
     ```
     Which in Anki becomes:
     
     <ruby>答<rt>こた</rt></ruby>えが<ruby>分<rt>わ</rt></ruby>かる


## How to use the template

### Default columns & configuration

#### Default constant values:
ROMAJI_COLUMN = 4 (column E)

JISHO_RESULT_LIMIT = 5

#### Columns:
![template screenshot](/imgs/template.png?raw=true "Template screenshot")

| Column | Description |
| --- | --- |
| Ruby | HTML Ruby tag created from Kanji (B) and Reading (C) columns |
| Kanji | Kanji parsed from selected dropdown (D) value |
| Reading | hiragana/katakana parsed from selected dropdown (D) value |
| Dropdown | a dropdown with top JISHO_RESULT_LIMIT results from Jisho search |
| Romaji | (**\***) romaji that acts as an input for Jisho search and dropdown population |
| LT meaning | (**\***) meaning of the word in my language |
| EN meaning | (**\***) english meaning of the word |
| Tags | concatenated tags from columns Tag1 - Tag5 |
| Tag1 - Tag5 | (**\***) anki card's tags |

(**\***) - **value is entered manually**

### Demo

![template demo](/gifs/demo.gif "Template demo")

1. Copy formulas to the next row
2. Enter romaji value
3. Select the desired search result from the dropdown
4. Enter the meanings and adjust tags


### Initial Setup

1. Install [LibreOffice GetRest Plugin](https://github.com/DmytroBazunov/LibreOfficeGetRestPlugin/wiki)

    As its binaries are now hard to come by, it is added in this repository under /plugins directory. **All credit goes to the creator**. Official plugin site can be found [here](https://extensions.libreoffice.org/en/extensions/show/libreoffice-getrest-plugin-1). To install, just double click the file and open it with LibreOffice.

2. Enable macros in LibreOffice Calc (`Tools > Options > LibreOffice > Security > Macro Security`)
3. Start using the template! For customizations, edit the constants in the jisho module (`Tools > Macros > Edit Macros` and look for modules under AnkiTemplate.ods)


## Importing to Anki

1. In LibreOffice Calc, go to File -> Save As...
2. Select `Text CSV` as the save type and click Save
3. In the dialog, choose `;` as a field delimeter
4. Click Ok
5. In Anki, go to File -> Import...
6. Choose the CSV file
7. Map CSV fields to your Anki Note Type's fields. Tags column (H) should be mapped to anki card's tags
8. Click Import

## Useful resources

1. [LibreOffice Getting Started With Macros](https://documentation.libreoffice.org/assets/Uploads/Documentation/en/GS5.1/HTML/GS5113-GettingStartedWithMacros.html#__RefHeading__5166_1196992793)
2. [Jisho API](https://jisho.org/forum/54fefc1f6e73340b1f160000-is-there-any-kind-of-search-api)
3. [How to import CSV into Anki](https://docs.ankiweb.net/importing.html#importing)
