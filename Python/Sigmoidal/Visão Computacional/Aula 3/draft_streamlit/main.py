# importar pacotes
import streamlit as st
import numpy as np
import cv2
from PIL import Image, ImageEnhance

OUTPUT_WIDTH = 500


def main():
    st.sidebar.header("Photomoidal")
    st.sidebar.info("100% em Python")
    st.sidebar.markdown("App para aplicar filtros em imagens, utilizando a bilioteca OpenCV.")

    # menu com oções de páginas
    opcoes_menu = ["Filtros", "Sobre"]
    escolha = st.sidebar.selectbox("Escolha uma opção", opcoes_menu)

    our_image = Image.open("empty.jpg")

    if escolha == "Filtros":
        st.title("Masterclass Visão Computacional")
        st.text("por Carlos Melo")
        st.markdown(f"""
                        ℹ️ Projeto da Masterclass Introdução à Visão Computacional, do Carlos Melo (Sigmoidal).
                        Para saber mais sobre a Masterclass, acesse [esta página](https://pay.hotmart.com/K44730436X?checkoutMode=10&bid=1608039415553).
                        """)

        # carregar e exibir imagem
        # our_image = cv2.imread(file_name)  ---> Não vai dar certo
        st.subheader("Carregar arquivo de imagem")
        image_file = st.file_uploader("Escolha uma imagem", type=['jpg', 'jpeg', 'png'])

        if image_file is not None:
            our_image = Image.open(image_file)
            st.sidebar.text("Imagem Original")
            st.sidebar.image(our_image, width=150)

        col1, col2 = st.beta_columns(2)

        # filtros que podem ser aplicados
        filtros = st.sidebar.radio("Filtros", ['Original', 'Grayscale', 'Desenho', 'Sépia', 'Blur',
                                               'Canny', 'Contraste'])

        if filtros == 'Grayscale':
            converted_image = np.array(our_image.convert('RGB'))
            gray_image = cv2.cvtColor(converted_image, cv2.COLOR_RGB2GRAY)

            col1.header("Original")
            col1.image(our_image, use_column_width=True)
            col2.header("Grayscale")
            col2.image(gray_image, use_column_width=True)
            # st.image(gray_image, width=OUTPUT_WIDTH, caption="Imagem com filtro Grayscale")

        elif filtros == 'Desenho':
            converted_image = np.array(our_image.convert('RGB'))
            gray_image = cv2.cvtColor(converted_image, cv2.COLOR_RGB2GRAY)
            inv_gray_image = 255 - gray_image
            blur_image = cv2.GaussianBlur(inv_gray_image, (21, 21), 0,0)
            sketch_image = cv2.divide(gray_image, 255 - blur_image, scale=256)

            col1.header("Original")
            col1.image(our_image, use_column_width=True)
            col2.header("Desenho")
            col2.image(sketch_image, use_column_width=True)
            # st.image(sketch_image, width=OUTPUT_WIDTH, caption="Imagem com filtro Desenho")

        elif filtros == 'Sépia':
            converted_image = np.array(our_image.convert('RGB'))
            converted_image = cv2.cvtColor(converted_image, cv2.COLOR_RGB2BGR)
            kernel = np.array([[0.272, 0.534, 0.131],
                               [0.349, 0.686, 0.168],
                               [0.393, 0.769, 0.189]])
            sepia_image = cv2.filter2D(converted_image, -1, kernel)

            col1.header("Original")
            col1.image(our_image, use_column_width=True)
            col2.header("Sépia")
            col2.image(sepia_image, channels="BGR", use_column_width=True)
            # st.image(sepia_image, channels="BGR", width=OUTPUT_WIDTH, caption="Imagem com filtro Sépia")

        elif filtros == 'Blur':
            b_amount = st.sidebar.slider("Kernel (n x n)", 3, 27, 9, step=2)
            converted_image = np.array(our_image.convert('RGB'))
            converted_image = cv2.cvtColor(converted_image, cv2.COLOR_RGB2BGR)
            blur_image = cv2.GaussianBlur(converted_image, (b_amount, b_amount), 0, 0)

            col1.header("Original")
            col1.image(our_image, use_column_width=True)
            col2.header("Sépia")
            col2.image(blur_image, channels="BGR", use_column_width=True)
            # st.image(blur_image, channels="BGR", width=OUTPUT_WIDTH, caption="Imagem com filtro Blur ({} x {}).".format(b_amount, b_amount))

        elif filtros == 'Canny':
            converted_image = np.array(our_image.convert('RGB'))
            converted_image = cv2.cvtColor(converted_image, cv2.COLOR_RGB2BGR)
            blur_image = cv2.GaussianBlur(converted_image, (11, 11), 0)
            canny = cv2.Canny(blur_image, 100, 150)

            col1.header("Original")
            col1.image(our_image, use_column_width=True)
            col2.header("Canny Edge Detection")
            col2.image(canny, use_column_width=True)
            # st.image(canny, width=OUTPUT_WIDTH, caption="Imagem com filtro Canny")

        elif filtros == "Contraste":
            c_amount = st.sidebar.slider("Constraste", 0.0, 2.0, 1.0)
            enhancer = ImageEnhance.Contrast(our_image)
            contrast_image = enhancer.enhance(c_amount)

            col1.header("Original")
            col1.image(our_image, use_column_width=True)
            col2.header("Contraste")
            col2.image(contrast_image, use_column_width=True)
            # st.image(contrast_image, width=OUTPUT_WIDTH, caption="Imagem com contraste em {}".format(c_amount))

        elif filtros == 'Original':
            st.image(our_image, width=OUTPUT_WIDTH)
        else:
            st.image(our_image, width=OUTPUT_WIDTH)

    elif escolha == 'Sobre':
        st.subheader("Este é um projeto da Masterclass Introdução à Visão Computacional.")
        st.markdown("Para saber mais informações, acesse [este site.](https://sigmoidal.ai)")
        st.text("Carlos Melo")
        st.success("Instagram @carlos_melo.py")
        st.video("https://www.youtube.com/watch?v=JhkhbTTxlQg")


if __name__ == '__main__':
    main()
