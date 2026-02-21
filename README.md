# hhxlsx2md
convert a vast xlsx with hierarchical header to a markdown for llm

# usage
```
$ python3 ./hhxlsx2md.py --geometry 3x4+2+3 sample.xlsx > sample.md
```

# concept
from:

|   | a |   |
| - | - | - |
|   | 1 | 2 |
| A | ! | @ |
| B | # | $ |

Sheet1

to:

# Sheet1
## A
- a
  - 1: !
  - 2: @
## B
- a
  - 1: #
  - 2: $
