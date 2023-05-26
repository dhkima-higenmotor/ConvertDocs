# ConvertDocs

## Commands

### PDF to PPTX

* https://github.com/ashafaei/pdf2pptx

```bash
pip install pdf2pptx
pdf2pptx test.pdf
```


### PNG to PPTX

* https://ggondae.tistory.com/42
* https://github.com/harsha2805/ImageToPptx

```bash
conda install python-pptx
```

```python
from pptx import Presentation
from pptx.util import Inches

prs = Presentation()

# Empty Layout
slide_layout = prs.slide_layout[6]
slide = prs.slide.add_slide(slide_layout)

left = Inches(1) # x좌표
top = Inches(1) # Y좌료
width = Inches(5) # 이미지 가로 길이
height = Inches(0.5) # 이미지 세로 길이

# 이미지 삽입
pic = slide.shape.add_pictures('test.png', ;eft, top, width, height)

prs.save('test,.pptx')
```



### PPTX to PDF

* 


## Python tools




