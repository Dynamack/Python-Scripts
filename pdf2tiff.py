from pdf2image import convert_from_path



## wrap everything in file iterator

## get dir
images = convert(r'C:\Users\om11\Documents\Project Blade\Biz Lease Agreements\Test\Allcap Limited - (Unit 24G) - 260521(127198679.1).pdf')

#images = convert_from_path(dir + filename)
## get filename

## translate to tiff
for i in range(2):
    images.save('page_' + str(i) + '.tiff', 'TIFF')