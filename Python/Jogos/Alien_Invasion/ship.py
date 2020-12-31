import pygame
import os

class Ship():

    def __init__(self, ai_settings, screen):
        """Inicializa a espaçonave e define sua posição inicial."""
        self.screen = screen

        #Carrega a imagem da espaçonave e obtém seu rect
        caminho = os.getcwd() + '\\Jogos\\Alien_Invasion\\imagens\\'
        self.image = pygame.image.load(caminho + 'ship.bmp')
        self.rect = self.image.get_rect()
        self.screen_rect = screen.get_rect()

        # Inicia cada nova espaçonave na parte inferior central da tela
        self.rect.centerx = self.screen_rect.centerx
        self.rect.bottom = self.screen_rect.bottom

        # Flag de movimento
        self.moving_right = False
        self.moving_left = False
        self.moving_up = False
        self.moving_down = False

    def update(self):
        #Atualiza a posição da espaçonave de acordo com as flags de movimento
        if self.moving_right:
            if (self.rect.centerx < 1135):
                self.rect.centerx += 1

        if self.moving_left:
            if (self.rect.centerx > 66):
                self.rect.centerx -= 1

        if self.moving_up:
            if (self.rect.centery > 50):
                self.rect.centery -= 1

        if self.moving_down:
            if (self.rect.centery < 550):
                self.rect.centery += 1

    def blitme(self):
        """Desenha a espaçonave em sua posição atual."""
        self.screen.blit(self.image, self.rect)

    def center_ship(self):
        """Centraliza a espaçonave na tela."""
        self.center = self.screen_rect.centerx