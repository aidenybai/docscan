# Docscan 

Docscan is a lightweight document scanner. It allows users to open up document types and return the information inside as strings via regex.

**Requirements**:
1. zipfile
2. io
3. re
4. XML

**Usage**:
*Note: fileName must be in the directory*
Example: DocuScan("C:\\Users\\You\\Desktop\\folder1\\test.pdf")
1. Instantiate `class Docscan('fileName')`.
2. use `print(variable.returnFileText())`
3. use `print(variable.executeRegex('regex here'))`
4. use `print(executeHeaderRegex('regex here'))`
5. use `print(executeFooterRegex('regex here'))`

**Methods**:
1. `returnFileText()` - Returns the text of a file.
2. `executeRegex(regexExpression)` - creates a list of all matching cases of regexExpression
3. `executeHeaderRegex(regularExpression)` - creates a list of all matching cases of regexExpression in the header XML.
4. `executeFooterRegex(regularExpression)` - creates a list of all matching cases of regexExpression in the Footer XML.
