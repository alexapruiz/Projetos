# importar nossas bibliotecas
from imutils.video import VideoStream
import numpy as np
import cv2
import time


# iniciar o meu streaming
vs = VideoStream(src=0).start()
time.sleep(2.0)

while True:
    # ler meus frames
    frame = vs.read()

    # etapa 1
    gray_image = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    # etapa 2
    inv_gray_image = 255 - gray_image
    # etapa 3
    blur_image = cv2.GaussianBlur(inv_gray_image, (21, 21), 0, 0)
    # etapa 4
    sketch_image = cv2.divide(gray_image, 255 - blur_image, scale=256)

    # exibir frame na tela
    cv2.imshow("Frame", sketch_image)
    key = cv2.waitKey(1) & 0xFF
    if key == ord('q'):
        break

# destruir as janelas e interromper a webcam
cv2.destroyAllWindows()
vs.stop()

