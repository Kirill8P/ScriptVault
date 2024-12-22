# Regex Example for Excel  
This VBA script demonstrates the use of regular expressions in Excel to extract specific patterns from text.  
## Features  
- Extracts the first matching HTML-like tags and their content from text (e.g., `<tag>content</tag>`).  
- Processes data from the first column of `Sheet1` and outputs results to the third column.  

## Requirements  
- Enable the **"Microsoft VBScript Regular Expressions 5.5"** library in the VBA editor:  
  1. Open the VBA editor (`Alt + F11`).  
  2. Go to `Tools` > `References`.  
  3. Check **"Microsoft VBScript Regular Expressions 5.5"**.  

## Usage  
1. Paste your data into the first column of `Sheet1`.  
2. Run the macro `regExExample`.  
3. Review the extracted content in the third column.

#### Input and Output:
| **Input**                                    | **Output**                              |
|----------------------------------------------|-----------------------------------------|
| `<sup>Superscript</sup>`                     | `<sup>Superscript</sup>`                |
| `<sub>Name</sub><div>Surname</div>`          | `<sub>Name</sub>; <div>Surname</div>`   |
| `<sub>Name<div>Surname</div></sub>`           | `<div>Surname</div>`                   |
| `<h1>Title<h1>`                              | `<h1>Title<h1>`                         |
| `<samp>Sample<samp>`                         | `<samp>Sample<samp>`                   |
| `<abbr>Abbreviation</abbr>`                  | `<abbr>Abbreviation</abbr>`            |
| `<del>Deleted<del>`                          | `<del>Deleted<del>`                    |
| `<div>Hello<div>`                            | `<div>Hello<div>`                      |
| `<div>Hello`                                 | `<div>Hello`                           |
| `<b><b>Bold</b>`                             | `<b>Bold</b>`                          |
| `<dfn>Definition</dfn>`                      | `<dfn>Definition</dfn>`                |
| `<sub>Subscript</sub>`                       | `<sub>Subscript</sub>`                 |
| `<a>Link</a>`                                | `<a>Link</a>`                          |
| `<strong>Strong`                             | `<strong>Strong`                       |
| `<code>Code</code>`                          | `<code>Code</code>`                    |

