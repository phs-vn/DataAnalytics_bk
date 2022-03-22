from automation.get_captcha import *

########################################################################################################
# Mr Hiep's code
imgPATH = r'C:\Users\namtran\Share Folder\Get Captcha\ivb_dataset\dataset\captcha_27.png'


def ivb_mr_hiep():
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
    return predictedCAPTCHA


########################################################################################################

# Nam's code

def ivb_captcha(i):
    input_path = r'C:\Users\namtran\Share Folder\Get Captcha\ivb_dataset\dataset'
    output_path = r'C:\Users\namtran\Share Folder\Get Captcha\ivb_dataset\fixed dataset'

    imagePATH = cv2.imread(join(input_path, fr'captcha_{i}.png'))
    gray = cv2.cvtColor(imagePATH, cv2.COLOR_BGR2GRAY)
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_OTSU + cv2.THRESH_BINARY_INV)[1]

    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (20, 1))
    detected_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, horizontal_kernel)

    cnts = cv2.findContours(detected_lines, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    cnts = cnts[0] if len(cnts) == 2 else cnts[1]

    for c in cnts:
        cv2.drawContours(imagePATH, [c], -1, (255, 255, 255), -1)

    repair_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 5))
    result = 255 - cv2.morphologyEx(255 - imagePATH, cv2.MORPH_CLOSE, repair_kernel, iterations=1)

    img = cv2.resize(result, (400,100))
    cv2.imwrite(join(output_path, fr'captcha_{i}.png'), img)

    predictedCAPTCHA = pytesseract.image_to_string(
        join(output_path, fr'captcha_{i}.png')
    )
    return predictedCAPTCHA


########################################################################################################

# Converting an image to black and white
def ivb_bw(i):
    original_path = r'C:\Users\namtran\Share Folder\Get Captcha\ivb_dataset\dataset'
    black_white_path = r'C:\Users\namtran\Share Folder\Get Captcha\ivb_dataset\black white dataset'

    originalImage = cv2.imread(join(original_path, fr'captcha_{i}.png'))
    grayImage = cv2.cvtColor(originalImage, cv2.COLOR_BGR2GRAY)

    (thresh, blackAndWhiteImage) = cv2.threshold(grayImage, 127, 255, cv2.THRESH_BINARY)

    cv2.imwrite(join(black_white_path, fr'captcha_{i}.png'), blackAndWhiteImage)

########################################################################################################
