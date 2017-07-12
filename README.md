# avbr_docx

avbr_docx is a fork of the [python-docx] module. 
The main purpose is to simplify the syntax of writing docx-files.
Sample usage:

```py
from avbr_docx import NewDocument, Text

doc = NewDocument("hello")
insertPoints = {}
doc.addText(Text("Hello", bold=True, italic=True))
insertPoints["point1"] = doc.getInsertPoint()
doc.addMultiText([Text("Hello ", bold= True, italic=True),
                  Text("World", bold =False, italic =True)])
#doc.addPicture("test.jpg")
doc.addText(Text("Hello", bold=True, italic=True),insertPoint=insertPoints["point1"])
doc.addHeading("Hello", level = 2)
doc.save()
```

   [python-docx]: <https://python-docx.readthedocs.io/en/latest/>

### Installation

See [python-docx] documentation for requirements.
Install the module by running "pip install ." in terminal/cmd.

```
cd directoryofpackage
pip install .
```


### Todos

 - Add more functionality


License
----

MIT

**Free Software, Hell Yeah!**
