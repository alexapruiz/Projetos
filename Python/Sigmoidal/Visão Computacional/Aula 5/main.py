# instalar
# pip install --upgrade google-cloud-vision

# MASTERCLASS
# VISÃO COMPUTACIONAL & OPENCV
# -- Sigmoidal

# importar as bibliotecas necessárias
import os
import cv2
import imutils
import numpy as np

# Credenciais da Google Vision API
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "credenciais/minhas_credenciais_da_api.json"


def preprocessing(path):
    """
    Pré-processa a imagem original, extraindo apenas a região da placa.

    :param path: Endereço da foto.
    :return: Recorte da placa do carro.
    """
    image = cv2.imread(path)
    cv2.imshow("image", image)
    cv2.waitKey(3000)

    # converter para tons de cinza
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    cv2.imshow("image", gray)
    cv2.waitKey(1000)
    # reduzir o nível de detalhes da foto
    gray = cv2.bilateralFilter(gray, 13, 15, 15)
    # cv2.imshow("image", gray)
    # cv2.waitKey(20000)
    # identificar bordas na foto
    edged = cv2.Canny(gray, 30, 200)
    cv2.imshow("image", edged)
    cv2.waitKey(1000)
    # identificar contornos
    contornos = cv2.findContours(edged.copy(), cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    contornos = imutils.grab_contours(contornos)
    contornos = sorted(contornos, key=cv2.contourArea, reverse=True)[:10]
    cv2.imshow("image", contornos)
    cv2.waitKey(1000)

    for c in contornos:
        peri = cv2.arcLength(c, True)
        approx = cv2.approxPolyDP(c, 0.018 * peri, True)
        break


preprocessing("000.jpg")
