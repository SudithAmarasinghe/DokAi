import cv2
import numpy as np
import pandas as pd
import pytesseract
import matplotlib.pyplot as plt
import statistics
import json
import os

class OCR:
    def __init__(self, image_folder, output_folder_json, output_folder_txt):
        self.image_folder = image_folder
        self.output_folder_json = output_folder_json
        self.output_folder_txt = output_folder_txt
        plt.rcParams['figure.figsize'] = [15, 8]

    def process_image(self, image_path, output_path_json, output_path_txt):
        img = cv2.imread(image_path, 0)
        img1 = cv2.copyMakeBorder(img, 50, 50, 50, 50, cv2.BORDER_CONSTANT, value=[255, 255])
        img123 = img1.copy()
        (thresh, th3) = cv2.threshold(img1, 128, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
        th3 = 255 - th3

        if th3.shape[0] < 1000:
            ver = np.array([[1], [1], [1], [1], [1], [1], [1]])
            hor = np.array([[1, 1, 1, 1, 1, 1]])
        else:
            ver = np.array([[1], [1], [1], [1], [1], [1], [1], [1], [1], [1], [1], [1], [1], [1], [1], [1], [1], [1], [1]])
            hor = np.array([[1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]])

        img_temp1 = cv2.erode(th3, ver, iterations=3)
        verticle_lines_img = cv2.dilate(img_temp1, ver, iterations=3)

        img_hor = cv2.erode(th3, hor, iterations=3)
        hor_lines_img = cv2.dilate(img_hor, hor, iterations=4)
        hor_ver = cv2.add(hor_lines_img, verticle_lines_img)
        hor_ver = 255 - hor_ver

        temp = cv2.subtract(th3, hor_ver)
        temp = 255 - temp
        tt = cv2.bitwise_xor(img1, temp)
        iii = cv2.bitwise_not(tt)
        tt1 = iii.copy()

        ver1 = np.array([[1, 1], [1, 1], [1, 1], [1, 1], [1, 1], [1, 1], [1, 1], [1, 1], [1, 1]])
        hor1 = np.array([[1, 1, 1, 1, 1, 1, 1, 1, 1, 1], [1, 1, 1, 1, 1, 1, 1, 1, 1, 1]])

        temp1 = cv2.erode(tt1, ver1, iterations=1)
        verticle_lines_img1 = cv2.dilate(temp1, ver1, iterations=1)

        temp12 = cv2.erode(tt1, hor1, iterations=1)
        hor_lines_img2 = cv2.dilate(temp12, hor1, iterations=1)
        hor_ver = cv2.add(hor_lines_img2, verticle_lines_img1)
        dim1 = (hor_ver.shape[1], hor_ver.shape[0])
        dim = (hor_ver.shape[1] * 2, hor_ver.shape[0] * 2)
        resized = cv2.resize(hor_ver, dim, interpolation=cv2.INTER_AREA)
        want = cv2.bitwise_not(resized)

        if want.shape[0] < 1000:
            kernel1 = np.array([[1, 1, 1]])
            kernel2 = np.array([[1, 1], [1, 1]])
            kernel3 = np.array([[1, 0, 1], [0, 1, 0], [1, 0, 1]])
        else:
            kernel1 = np.array([[1, 1, 1, 1, 1, 1]])
            kernel2 = np.array([[1, 1, 1, 1, 1], [1, 1, 1, 1, 1], [1, 1, 1, 1, 1], [1, 1, 1, 1, 1]])

        tt1 = cv2.dilate(want, kernel1, iterations=14)
        resized1 = cv2.resize(tt1, dim1, interpolation=cv2.INTER_AREA)
        contours1, hierarchy1 = cv2.findContours(resized1, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)

        def sort_contours(cnts, method="left-to-right"):
            reverse = False
            i = 0
            if method == "right-to-left" or method == "bottom-to-top":
                reverse = True
            if method == "top-to-bottom" or method == "bottom-to-top":
                i = 1
            boundingBoxes = [cv2.boundingRect(c) for c in cnts]
            (cnts, boundingBoxes) = zip(*sorted(zip(cnts, boundingBoxes), key=lambda b: b[1][i], reverse=reverse))
            return (cnts, boundingBoxes)

        (cnts, boundingBoxes) = sort_contours(contours1, method="top-to-bottom")
        heightlist = []
        for i in range(len(boundingBoxes)):
            heightlist.append(boundingBoxes[i][3])
        heightlist.sort()
        sportion = int(.5 * len(heightlist))
        eportion = int(0.05 * len(heightlist))
        try:
            medianheight = statistics.mean(heightlist[-sportion:-eportion])
        except:
            medianheight = 0

        box = []
        imag = iii.copy()
        for i in range(len(cnts)):
            cnt = cnts[i]
            x, y, w, h = cv2.boundingRect(cnt)
            if h >= .7 * medianheight and w / h > 0.9:
                image = cv2.rectangle(imag, (x + 4, y - 2), (x + w - 5, y + h), (0, 255, 0), 1)
                box.append([x, y, w, h])

        main = []
        j = 0
        l = []
        for i in range(len(box)):
            if i == 0:
                l.append(box[i])
                last = box[i]
            else:
                if box[i][1] <= last[1] + medianheight / 2:
                    l.append(box[i])
                    last = box[i]
                    if i == len(box) - 1:
                        main.append(l)
                else:
                    main.append(l)
                    l = []
                    last = box[i]
                    l.append(box[i])

        maxsize = 0
        for i in range(len(main)):
            l = len(main[i])
            if maxsize <= l:
                maxsize = l

        ylist = []
        for i in range(len(boundingBoxes)):
            ylist.append(boundingBoxes[i][0])

        ymax = max(ylist)
        ymin = min(ylist)

        ymaxwidth = 0
        for i in range(len(boundingBoxes)):
            if boundingBoxes[i][0] == ymax:
                ymaxwidth = boundingBoxes[i][2]

        TotWidth = ymax + ymaxwidth - ymin
        width = []
        widthsum = 0
        for i in range(len(main)):
            for j in range(len(main[i])):
                widthsum = main[i][j][2] + widthsum
            width.append(widthsum)
            widthsum = 0

        main1 = main.copy()
        maxsize1 = 0
        for i in range(len(main1)):
            l = len(main1[i])
            if maxsize1 <= l:
                maxsize1 = l

        midpoint = []
        for i in range(len(main1)):
            if len(main1[i]) == maxsize1:
                for j in range(maxsize1):
                    midpoint.append(int(main1[i][j][0] + main1[i][j][2] / 2))
                break

        midpoint = np.array(midpoint)
        midpoint.sort()
        final = [[] * maxsize1] * len(main1)

        for i in range(len(main1)):
            for j in range(len(main1[i])):
                min_idx = j
                for k in range(j + 1, len(main1[i])):
                    if main1[i][min_idx][0] > main1[i][k][0]:
                        min_idx = k
                main1[i][j], main1[i][min_idx] = main1[i][min_idx], main1[i][j]

        finallist = []
        for i in range(len(main1)):
            lis = [[] for k in range(maxsize1)]
            for j in range(len(main1[i])):
                diff = abs(midpoint - (main1[i][j][0] + main1[i][j][2] / 4))
                minvalue = min(diff)
                ind = list(diff).index(minvalue)
                lis[ind].append(main1[i][j])
            finallist.append(lis)

        todump = []
        for i in range(len(finallist)):
            for j in range(len(finallist[i])):
                to_out = ''
                if len(finallist[i][j]) == 0:
                    todump.append(' ')
                else:
                    for k in range(len(finallist[i][j])):
                        y, x, w, h = finallist[i][j][k][0], finallist[i][j][k][1], finallist[i][j][k][2], \
                                     finallist[i][j][k][3]
                        roi = iii[x:x + h, y + 2:y + w]
                        roi1 = cv2.copyMakeBorder(roi, 5, 5, 5, 5, cv2.BORDER_CONSTANT, value=[255, 255])
                        img = cv2.resize(roi1, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
                        kernel = np.ones((2, 1), np.uint8)
                        img = cv2.dilate(img, kernel, iterations=1)
                        img = cv2.erode(img, kernel, iterations=2)
                        img = cv2.dilate(img, kernel, iterations=1)
                        out = pytesseract.image_to_string(img)
                        if len(out) == 0:
                            out = pytesseract.image_to_string(img)
                        to_out += " " + out
                    todump.append(to_out)
        npdump = np.array(todump)
        dataframe = pd.DataFrame(npdump.reshape(len(main1), maxsize1))

        with open(output_path_json, 'w') as json_file:
            json_str = dataframe.to_json(orient='index')
            json_obj = json.loads(json_str)
            json.dump(json_obj, json_file)

        npdump = np.array(todump)
        dataframe = pd.DataFrame(npdump.reshape(len(main1), maxsize1))

        with open(output_path_txt, 'w') as txt_file:
            for row in dataframe.itertuples(index=False, name=None):
                txt_file.write('\t'.join(map(str, row)) + '\n')

    def process_directory(self):
        existing_files_json = os.listdir(self.output_folder_json)
        for file_json in existing_files_json:
            file_path_json = os.path.join(self.output_folder_json, file_json)
            if os.path.isfile(file_path_json):
                os.remove(file_path_json)

        existing_files_txt = os.listdir(self.output_folder_txt)
        for file_txt in existing_files_txt:
            file_path_txt = os.path.join(self.output_folder_txt, file_txt)
            if os.path.isfile(file_path_txt):
                os.remove(file_path_txt)

        os.makedirs(self.output_folder_json, exist_ok=True)
        os.makedirs(self.output_folder_txt, exist_ok=True)

        files = os.listdir(self.image_folder)
        image_files = [file for file in files if file.endswith(('.jpg', '.jpeg', '.png', '.bmp'))]

        for image_file in image_files:
            image_path = os.path.join(self.image_folder, image_file)
            output_file_json = os.path.splitext(image_file)[0] + '.json'
            output_file_txt = os.path.splitext(image_file)[0] + '.txt'
            output_path_json = os.path.join(self.output_folder_json, output_file_json)
            output_path_txt = os.path.join(self.output_folder_txt, output_file_txt)
            self.process_image(image_path, output_path_json, output_path_txt)

    def run(self):
        
        self.process_directory()
        



