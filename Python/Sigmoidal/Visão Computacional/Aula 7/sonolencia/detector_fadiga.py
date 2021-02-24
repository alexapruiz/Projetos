# Inspirado no projeto do mestre da Visão Computacional, Adrian Rosebrock
# https://bit.ly/2CYC7Gf

# importar pacotes necessários
from scipy.spatial import distance as dist
from imutils.video import VideoStream
from imutils import face_utils
from threading import Thread
import numpy as np
import playsound
import imutils
import time
import dlib
import cv2
import matplotlib.pyplot as plt


# definir constantes
ALARME_SONORO = "buzina.wav"
WEBCAM = 0
OLHO_THRESHOLD = 0.25
NUM_FRAMES_CONSEC = 100
CONTADOR = 0
ALARME_LIGADO = False


def disparar_alarme(path=ALARME_SONORO):
    """
    Dispara um alarme sonoro a partir de um arquivo .wav

    :param path: localização do arquivo wav
    :return: None
    """

    playsound.playsound(ALARME_SONORO)
    return None


def calcular_ratio_olho(eye):
    """
    Calcula o aspecto (ratio) do olho.
    Fonte: https://towardsdatascience.com/drowsiness-detection-system-in-real-time-using-opencv-and-flask-in-python-b57f4f1fcb9e

    :param eye:
    :return:
    """
    # compute the euclidean distances between the two sets of
    # vertical eye landmarks (x, y)-coordinates
    A = dist.euclidean(eye[1], eye[5])
    B = dist.euclidean(eye[2], eye[4])

    # compute the euclidean distance between the horizontal
    # eye landmark (x, y)-coordinates
    C = dist.euclidean(eye[0], eye[3])

    # compute the eye aspect ratio
    ear = (A + B) / (2.0 * C)

    # return the eye aspect ratio
    return ear


# carregar o dlib para face detector
detector = dlib.get_frontal_face_detector()
predictor = dlib.shape_predictor("shape_predictor_68_face_landmarks.dat")

# pegar os índices do previsor, para olhos esquerdo e direito
(lStart, lEnd) = face_utils.FACIAL_LANDMARKS_IDXS["left_eye"]
(rStart, rEnd) = face_utils.FACIAL_LANDMARKS_IDXS["right_eye"]

# inicializar vídeo
print("[INFO] inicializando streaming de vídeo...")
vs = VideoStream(src=WEBCAM).start()
time.sleep(2.0)

# Tirar o comentário
# # desenhar um objeto do tipo figure
# y = [None] * 100
# x = np.arange(0,100)
# fig = plt.figure(figsize=(5,3))
# ax = fig.add_subplot(111)
# li, = ax.plot(x, y)

# loop sobre os frames do vídeo
while True:
    frame = vs.read()
    frame = imutils.resize(frame, width=800)
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

    # detectar faces (grayscale)
    rects = detector(gray, 0)

    # loop nas detecções de faces
    for rect in rects:
        shape = predictor(gray, rect)
        shape = face_utils.shape_to_np(shape)

        # extrair coordenadas dos olhos e calcular a proporção de abertura
        leftEye = shape[lStart:lEnd]
        rightEye = shape[rStart:rEnd]
        leftEAR = calcular_ratio_olho(leftEye)
        rightEAR = calcular_ratio_olho(rightEye)

        # ratio médio para os dois olhos
        ear = (leftEAR + rightEAR) / 2.0

        # convex hull para os olhos
        leftEyeHull = cv2.convexHull(leftEye)
        rightEyeHull = cv2.convexHull(rightEye)
        cv2.drawContours(frame, [leftEyeHull], -1, (0, 255, 0), 1)
        cv2.drawContours(frame, [rightEyeHull], -1, (0, 255, 0), 1)

        # # salvar historico para plot
        # y.pop(0)
        # y.append(ear)

        # # update canvas
        # plt.xlim([0, 100])
        # plt.ylim([0, 0.4])
        # ax.relim()
        # ax.autoscale_view(True, True, True)
        # fig.canvas.draw()
        # plt.show(block=False)
        # li.set_ydata(y)
        # fig.canvas.draw()
        # time.sleep(0.01)

        # checar ratio x threshold
        if ear < OLHO_THRESHOLD:
            CONTADOR += 1

            # dentro dos critérios, soar o alarme
            if CONTADOR >= NUM_FRAMES_CONSEC:
                # ligar alarme
                if not ALARME_LIGADO:
                    ALARME_LIGADO = True
                    t = Thread(target=disparar_alarme)
                    t.deamon = True
                    t.start()

                cv2.putText(frame, "[ALERTA] FADIGA!", (10, 30),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 255), 2)

        # caso acima do threshold, resetar o contador e desligar o alarme
        else:
            CONTADOR = 0
            ALARME_LIGADO = False

            # desenhar a proporção de abertura dos olhos
        cv2.putText(frame, "EAR: {:.2f}".format(ear), (300, 30),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 255), 2)

    # mostrar o frame
    cv2.imshow("Frame", frame)
    key = cv2.waitKey(1) & 0xFF

    # tecla para sair do script "q"
    if key == ord("q"):
        break

# clean
cv2.destroyAllWindows()
vs.stop()
