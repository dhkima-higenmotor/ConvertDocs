import aspose.slides as slides

with slides.Presentation() as pres:
    pres.slides.remove_at(0)
    pres.slides.add_from_pdf("test2.pdf")
    pres.save("test2.pptx", slides.export.SaveFormat.PPTX)
