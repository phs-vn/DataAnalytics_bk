from automation import *
# img = cv2.imread(r'C:\Users\namtran\Share Folder\Get Captcha\ivb_dataset\captcha_1.png')
imgPATH = r'C:\Users\namtran\Share Folder\Get Captcha\ivb_dataset\captcha_1.png'
img = Image.open(imgPATH)
data = np.array(img)
data = data[:, :, :3]
data[np.sum(data, axis=2) > 400] = 0
img = Image.fromarray(data)
img = img.resize((720, 200))
# img.save(join(dirname(__file__), 'CAPTCHA', f'IVB_bk.png'))
img.save(join(r'C:\Users\namtran\Share Folder\Get Captcha\ivb_fixed\IVB_bk.png'))
# predictedCAPTCHA = pytesseract.image_to_string(join(dirname(__file__), 'CAPTCHA', f'IVB_bk.png'))

########################################################################################################

