from textblob import TextBlob

blob = TextBlob('That movie was good')
print(blob.subjectivity,blob.polarity)
print('===============================================')

blob = TextBlob('That movie was awesome')
print(blob.subjectivity,blob.polarity)