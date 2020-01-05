# python-excel-report-generator
python-excel-report-generator is a microservice written in python to write an excel report in xlsx file. It is born from the lack of an existing library to write natively from Python the Office Open XML format. It can be accessed from any other programming framework using post request. 

### Framework and languages
* Django 2.2
* python 3.6
* pandas
* Django Rest Framework
* openpyxl

### How to use this service
**url = "http://excel.iofact.com/api/excel_export"**

**request method = post**

**data type = json**

#### Data preparation
**"topHeader"**: topHeader is a list of dictionary. In topHeader, user must explicitly define the cell information with column index. 

```python
topHeader = [{
    "column": "A11:A13",
    "title": "Cell value",
    "font": {
        "font_size": "11",
        'font_family': "Calibri",
        "bold": True,
        "italic": False,
        "underline": "none",
        "color": "FF000000"
    },
    "alignment": {
        "horizontal": "center",
        "vertical": "center"
    }
}]

```
**"columnHeader"**: columnHeader contains the column information of the excel report. Users can create a dynamic excel report with the option of cell merging using columnHeader. In this key, there is no need to tell the excel column index. Insted children cell (which cell need to be merged) should send in children parameter. However, a key parameter should be sent for mapping the data with the column.  

```python
columnHeader = [
    {'title': 'A', 'key': 'a', 'style': {'font': {'font_size': '11', 'font_family': 'Calibri', 'bold': True, 'italic': False,
                      'underline': 'none', 'color': 'FF000000'},
             'alignment': {'horizontal': 'center', 'vertical': 'center'}}},
    {'title': 'B', 'key': 'b',
     'children':
         [
             {'title': 'C', 'key': 'c'},
             {'title': 'X', 'key': 'x'},
             {'title': 'D', 'key': 'd',
              'children': [
                  {'title': 'E', 'key': 'e'},
                  {'title': 'F', 'key': 'f'}
              ]
              }
         ],
     },
    {
        'title': 'G', 'key': 'g',
        'children': [
             {'title': 'H', 'key': 'h'},
             {'title': 'J', 'key': 'j'},
             {'title': 'I', 'key': 'i',
              'children': [
                  {'title': 'K', 'key': 'k'},
                  {'title': 'L', 'key': 'l',
                   'children': [
                        {'title': 'H', 'key': 'h'},
                        {'title': 'J', 'key': 'j'}]
                   }
              ]
              }
        ]
    },
    {
        'title': 'Z', 'key': 'z'
    }
]
```

**"explicitColumnHeader"**: Explicitly define the cell index. No children paremeter will be accepted. Used method is as same as "topHeader."

```python
explicitColumnHeader = [{
  "column": "A1:A3",
  "title": "Cell information"
}]
```
**"tableData"**: All the cell data will be send in "tableData". 

#### Sample Data preparation (API testing)
* simple Api Data
```json
data = { 
   "explicitColumnHeader":[ 
      { 
         "column":"A1:A3",
         "title":"Budget"
      },
      { 
         "column":"B1:B3",
         "title":"Events"
      }
   ],
   "tableData":{ 
      "Budget":[ 
         10000,
         15000,
         20000
      ],
      "Events":[ 
         "A",
         "B",
         "C"
      ]
   }
}

excelReport = requests.post("http://excel.iofact.com/api/excel_export", json=data)
```

### How to use this service
The service accepts a JSON object with two keys from the post request. The first key, either "columnHeader" or "explicitColumnHeader" will be a list of json object or dictionary. The "explicitColumnHeader" must contain the requirements of the excel report such as cell information, alignment, font. A cell information and its requirment can spcify by the following way:
```python
[{
  "column": "A1:A3",
  "title": "Cell information"
}]
```
'column' A1:A3 will merge the column 1 to 3 of cell A in the excel and then the title value 'Cell information' will kept in this cell. Some more paremeter could be passed in the header. The following code will show how to adjust font and alignment of a cell. By default, font size is 11, font family is Calibri, Boldface, Italic and underline is false, and color is black. 

```python
header = [{
    "column": "A11:A13",
    "title": "Cell value",
    "font": {
        "font_size": "11",
        'font_family': "Calibri",
        "bold": True,
        "italic": False,
        "underline": "none",
        "color": "FF000000"
    },
    "alignment": {
        "horizontal": "center",
        "vertical": "center"
    }
}]
```
The "columnHeader" offers a dynamic cell creation method where user can send their object in a tree structure. This can be created in the following way:

```python
head = [
    {'title': 'A', 'key': 'a', 'style': {'font': {'font_size': '11', 'font_family': 'Calibri', 'bold': True, 'italic': False,
                      'underline': 'none', 'color': 'FF000000'},
             'alignment': {'horizontal': 'center', 'vertical': 'center'}}},
    {'title': 'B', 'key': 'b',
     'children':
         [
             {'title': 'C', 'key': 'c'},
             {'title': 'X', 'key': 'x'},
             {'title': 'D', 'key': 'd',
              'children': [
                  {'title': 'E', 'key': 'e'},
                  {'title': 'F', 'key': 'f'}
              ]
              }
         ],
     },
    {
        'title': 'G', 'key': 'g',
        'children': [
             {'title': 'H', 'key': 'h'},
             {'title': 'J', 'key': 'j'},
             {'title': 'I', 'key': 'i',
              'children': [
                  {'title': 'K', 'key': 'k'},
                  {'title': 'L', 'key': 'l',
                   'children': [
                        {'title': 'H', 'key': 'h'},
                        {'title': 'J', 'key': 'j'}]
                   }
              ]
              }
        ]
    },
    {
        'title': 'Z', 'key': 'z'
    }
]
```
The algorithm will automatically create the cell number for this formate. In this method, a key is required for map the data with it's key. 


The second key is "df," where the data will send in the form of JSON, dictionary, or data frame object. A json object of dataframe could be created by the following rules:

```python
dfJson = [{
        "Title": 'Project Introduction',
        "Target": 100,
        "Acheive": 90
    },
    {
        "Title": "Project Organization",
        "Target": 100,
        "Acheive": 90
    },
]
```
If the dataframe will send as json object, than it shoulb be dumps in json from python. It is required to import json library to dumps in json object.
```python
import json
JsonDf = json.dumps(dfJson)
```
Now the data could be prepare for the api request. 

**url = "http://excel.iofact.com/api/excel_export"**

**request method = post**

**data type = json**

```python
excelReport = requests.post("http://excel.iofact.com/api/excel_export", json={"explicitColumnHeader": header, "df": JsonDf})
```


#### Sample Data preparation (API testing)
* simple Api Data
```json
data = { 
   "header":[ 
      { 
         "column":"A1:A3",
         "title":"Budget"
      },
      { 
         "column":"B1:B3",
         "title":"Events"
      }
   ],
   "df":{ 
      "Budget":[ 
         10000,
         15000,
         20000
      ],
      "Events":[ 
         "A",
         "B",
         "C"
      ]
   }
}

excelReport = requests.post("http://excel.iofact.com/api/excel_export", json=data)
```

### Sample Code Example in python
* Step 1: Preparing the data (header and df)

```python
header = [{
        "column": "A1:A2",
        "title": "Col 1",
        "font": {
            "font_size": "11",
            'font_family': "Calibri",
            "bold": True,
            "italic": False,
            "underline": "none",
            "color": "FF000000"
        },
        "alignment": {
            "horizontal": "center",
            "vertical": "center"
        }
    },
    {
        "column": "B1:C1",
        "title": "Col 2",
        "font": {
            "font_size": "11",
            'font_family': "Calibri",
            "bold": True,
            "italic": False,
            "underline": "none",
            "color": "FF000000"
        },
        "alignment": {
            "horizontal": "center",
            "vertical": "center"
        }
    },
    {
        "column": "B2:B2",
        "title": "Col 2.1",
        "font": {
            "font_size": "11",
            'font_family': "Calibri",
            "bold": True,
            "italic": False,
            "underline": "none",
            "color": "FF000000"
        },
        "alignment": {
            "horizontal": "center",
            "vertical": "center"
        }
    },
    {
        "column": "C2:C2",
        "title": "Col 2.2",
        "font": {
            "font_size": "11",
            'font_family': "Calibri",
            "bold": True,
            "italic": False,
            "underline": "none",
            "color": "FF000000"
        },
        "alignment": {
            "horizontal": "center",
            "vertical": "center"
        }
    }
]

df = {
    "col 1": [15, 16, 17, 18],
    "col 2.1": [25, 26, 27, 28],
    "col 2.2": [33, 34, 35, 36],
}

data = {
    "header": header,
    "df": df
}
```

* Step 2: Api call for excel preparation

```python

import requests

excelReport = requests.post("http://excel.iofact.com/api/excel_export", json = data)

with open('/home/roomey/report.xlsx', 'wb') as f:
    f.write(excelReport.content)

f.close()

```
![alt text][sample output]

[sample output]: https://github.com/devnesthq/python-excel-report-generator/blob/master/Example/Sample%20Output
