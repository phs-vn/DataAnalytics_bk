from automation import *

########################################################################################################
# Mr Hiep's code
imgPATH = r'C:\Users\namtran\Share Folder\Get Captcha\ivb_dataset\dataset\captcha_27.png'
img = Image.open(imgPATH)
data = np.array(img)
data = data[:, :, :3]
data[np.sum(data, axis=2) > 400] = 0
img = Image.fromarray(data)
img = img.resize((720, 200))
img.save(join(r'C:\Users\namtran\Share Folder\Get Captcha\ivb_dataset\fixed dataset\captcha_27.png'))
predictedCAPTCHA = pytesseract.image_to_string(
    r'C:\Users\namtran\Share Folder\Get Captcha\ivb_dataset\fixed dataset\captcha_27.png'
)

########################################################################################################
# Nam's code
image = cv2.imread(r'C:\Users\namtran\Share Folder\Get Captcha\ivb_dataset\dataset\captcha_27.png')

gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_OTSU | cv2.THRESH_BINARY_INV)[1]

horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (20, 1))
detected_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, horizontal_kernel)

cnts = cv2.findContours(detected_lines, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
cnts = cnts[0] if len(cnts) == 2 else cnts[1]

for c in cnts:
   cv2.drawContours(image, [c], -1, (255,255,255), -1)

repair_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 5))
result = 255 - cv2.morphologyEx(255 - image, cv2.MORPH_CLOSE, repair_kernel,iterations=1)

img = Image.fromarray(result)
img = img.resize((400,100))
img.save(join(r'C:\Users\namtran\Share Folder\Get Captcha\ivb_dataset\fixed dataset\captcha_27.png'))

predictedCAPTCHA = pytesseract.image_to_string(
    r'C:\Users\namtran\Share Folder\Get Captcha\ivb_dataset\fixed dataset\captcha_27.png'
)

########################################################################################################
