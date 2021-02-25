from imutils.video import VideoStream
import imutils
import numpy as np
import cv2
import time
import argparse

# constantes
PROTOTXT = "deploy.prototxt.txt"
MODEL = "res10_300x300_ssd_iter_140000.caffemodel"
CONFIDENCE_THRESHOLD = 0.7
pos = (200, 200)

def anonymize_face_pixelate(image, blocks=9):
    """
    Função extraída do autor Adrian Rosebrock do site PyImageSearch
    https://www.pyimagesearch.com
    """

    # divide the input image into NxN blocks
    (h, w) = image.shape[:2]
    xSteps = np.linspace(0, w, blocks + 1, dtype="int")
    ySteps = np.linspace(0, h, blocks + 1, dtype="int")
    # loop over the blocks in both the x and y direction
    for i in range(1, len(ySteps)):
        for j in range(1, len(xSteps)):
            # compute the starting and ending (x, y)-coordinates
            # for the current block
            startX = xSteps[j - 1]
            startY = ySteps[i - 1]
            endX = xSteps[j]
            endY = ySteps[i]
            # extract the ROI using NumPy array slicing, compute the
            # mean of the ROI, and then draw a rectangle with the
            # mean RGB values over the ROI in the original image
            roi = image[startY:endY, startX:endX]
            (B, G, R) = [int(x) for x in cv2.mean(roi)[:3]]
            cv2.rectangle(image, (startX, startY), (endX, endY),
                          (B, G, R), -1)
    # return the pixelated blurred image
    return image

# carregar modelo
net = cv2.dnn.readNetFromCaffe(PROTOTXT, MODEL)

# inicializa o streaming
vs = VideoStream(src=0).start()
time.sleep(2.0)

while True:
    frame = vs.read()
    frame = imutils.resize(frame, width=400)

    h, w = frame.shape[:2]
    blob = cv2.dnn.blobFromImage(frame, 1.0, (300, 300), (104.0, 177.0, 123.0))
    net.setInput(blob)
    detections = net.forward()

    # iterar ao longo das deteccoes
    for i in range(0, detections.shape[2]):
        # exemplo de intervalo de confianca
        confidence = detections[0, 0, i, 2]

        # selecionar apenas intervalos acima do threshold
        if confidence > CONFIDENCE_THRESHOLD:
            # label da confian'ca
            text = "{:.2f}%".format(confidence * 100)

            # calcular o bounding box
            box = detections[0, 0, i, 3:7] * np.array([w, h, w, h])
            (startX, startY, endX, endY) = box.astype("int")

            kW = int(w / 3.0)
            kH = int(h / 3.0)
            face = frame[startY:endY, startX:endX]

            # blured = cv2.GaussianBlur(face, (kW, kH), 0)
            blured = anonymize_face_pixelate(face)

            frame[startY:endY, startX:endX] = blured

    cv2.imshow("Frame", frame)

    key = cv2.waitKey(1) & 0xFF
    if key == ord("q"):
        break

cv2.destroyAllWindows()
vs.stop()
