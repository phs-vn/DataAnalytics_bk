from request.stock import *


# CÁCH 1
img = cv2.imread(fr'D:\DataAnalytics\news_analysis\captcha data\captcha_1.jpg')
gry = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
(h, w) = gry.shape[:2]
gry = cv2.resize(gry, (w * 2, h * 2))
cls = cv2.morphologyEx(gry, cv2.MORPH_CLOSE, None)
thr = cv2.threshold(cls, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
txt = pytesseract.image_to_string(thr)
print(txt)

# CÁCH 2
img = r'D:\DataAnalytics\news_analysis\captcha data\captcha_1.jpg'


def read_image(img_path):
    try:
        return pytesseract.image_to_string(img_path).replace('\n', '')
    except:
        return "[ERROR] Unable to process file: {0}".format(img_path)
