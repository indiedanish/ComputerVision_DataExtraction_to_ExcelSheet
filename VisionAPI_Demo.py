import os, io
from google.cloud import vision_v1
from google.cloud.vision_v1 import types
import pandas as pd
import xlwings as xw
from xlwings import Range, constants

os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = r"ServiceAccountToken.json"

client = vision_v1.ImageAnnotatorClient()

file_name = 'img.jpeg'
image_path = r'E:\Clone\pythin\VisionAPI\Images'

with io.open(os.path.join(image_path, 'img.jpeg'),'rb') as image_file:
    content = image_file.read()

# construct an iamge instance
image = vision_v1.types.Image(content=content)

"""
# or we can pass the image url
image = vision.types.Image()
image.source.image_uri = 'https://edu.pngfacts.com/uploads/1/1/3/2/11320972/grade-10-english_orig.png'
"""

# annotate Image Response
response = client.text_detection(image=image)  # returns TextAnnotation


df = pd.DataFrame(columns=['locale', 'description'])

texts = response.text_annotations
for text in texts:
    df = df.append(
        dict(
            locale=text.locale,
            description=text.description
        ),
        ignore_index=True
    )

a = df['description'][0]
print(a.split('\n')[3:])

print('\n')

num_list = []

for num in a.split('\n')[3:]:
    if "+" in num:
        
        
        num = num.replace("+", "")
        num = num.replace(" ", "")
        num_list.append(num)

print(num_list)


wb = xw.Book(r'test.xlsm')
sheet = wb.sheets[0]
for index, element in enumerate(num_list) :
    sheet.range('A'+str(index+2)).expand('table').value  = element
    


wb.save()
