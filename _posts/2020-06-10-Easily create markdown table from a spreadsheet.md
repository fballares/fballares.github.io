---
layout: post
section-type: post
title:  "Easily create markdown table from a spreadsheet"
date:   2020-06-10
categories: tech
tags: [ 'python', 'code-snippets', 'markdown' ]
---

![markdown_table](/gallery/markdown_table.gif)


Markdown is really great, I like it for many good reasons.  This site is completely written in markdown.  Just like with any tool though, there will be minor quirks that we have to deal with.  For me, it's working with tables in markdown (it's really inconvenient).  My use-case is very simple.  I would like to easily convert a simple table I have on a spreadsheet and convert to markdown for note taking or blog post.  With my script below, I'm able to copy a table from a spreadsheet and then paste it as a markdown table format.  

See below:

``` python
# Load the modules required
# Using pyperclip to handle clipboard
# Using the pandas.read_clipboard function to load copied text from spreadsheets.
# Tabulate module is required for markdown to properly function (needed)

import pyperclip, tabulate
import pandas as pd

# Load what is in the clipboard.
# I'm assuming here that a table was copied from Excel/CSV file.
raw_table = pd.read_clipboard()

# convert to a dataframe
df = pd.DataFrame(data=raw_table)

# create the markdown output and put it into the clipboard.  
#After this, you can paste the clipboard content directly into your editor.

pyperclip.copy(df.to_markdown())

```

Here is the markdown output.

|    | Column A   | Column B   | Column C   |
|---:|:-----------|:-----------|:-----------|
|  0 | Data A1    | Data B1    | Data C1    |
|  1 | Data A2    | Data B2    | Data C2    |
|  2 | Data A3    | Data B3    | Data C3    |


[Link to the Gist](https://gist.github.com/fballares/7a3dcd0e884c5de81553bc162931c5c0)