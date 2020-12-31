import sys, pygame, os

pygame.init()

size = width, height = 600, 400
screen = pygame.display.set_mode(size)
black = 0, 0, 0

class Bouncer(pygame.sprite.Sprite):
    """classe para o bouncer"""

    def __init__(self, startpos):
        pygame.sprite.Sprite.__init__(self)
        self.direction = 1
        self.image, self.rect = load_image('bouncer.gif')
        self.rect.centerx = startpos[0]
        self.rect.centery = startpos[1]

    def update(self):
        self.rect.move_ip((self.direction * 3, 0))
        if self.rect.left < 0:
            self.direction = 1
        elif self.rect.right > width:
            self.direction = -1


class Ball(pygame.sprite.Sprite):
    """classe para a bola"""

    def __init__(self, startpos):
        pygame.sprite.Sprite.__init__(self)
        self.speed = [2, 2]
        self.image, self.rect = load_image('ball.gif')
        self.rect.centerx = startpos[0]
        self.rect.centery = startpos[1]
        self.init_pos = startpos

    def update(self):
        self.rect.move_ip(self.speed)
        if self.rect.left < 0 or self.rect.right > width:
            self.speed[0] = -self.speed[0]
        if self.rect.top < 0:
            self.speed[1] = -self.speed[1]
        if self.rect.bottom > height:
            self.rect.centerx = self.init_pos[0]
            self.rect.centery = self.init_pos[1]

def load_image(name):
    """carrega uma imagem na memoria, e retorna a imagem e o seu rect (retangulo)"""
    #fullname = os.path.join('images', name)
    fullname = os.getcwd() + '\\Forms\\images\\'
    image = pygame.image.load(fullname + str(name))
    print('image : ' + fullname + str(name))
    return image, image.get_rect()

def main():
    # cria os nossos objetos (bola e bouncer)
    ball = Ball([100, 100])
    bouncer = Bouncer([20, 395])
    pygame.display.set_caption('Bouncer!')
    clock = pygame.time.Clock()

    while 1:
        # garante que o programa nao vai rodar a mais que 120fps
        clock.tick(120)

        # checa eventos de teclado
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                sys.exit()
            elif event.type == pygame.KEYDOWN:
                if event.key == pygame.K_LEFT:
                    bouncer.direction = -1
                if event.key == pygame.K_RIGHT:
                    bouncer.direction = 1

        # checa se a bola colidiu no bouncer, e caso sim inverte a direcao vertical da bola
        if bouncer.rect.colliderect(ball.rect):
            if ball.speed[1] > 0:
                ball.speed[1] = -ball.speed[1]

        # atualiza os objetos
        ball.update()
        bouncer.update()

        # redesenha a tela
        screen.fill(black)
        screen.blit(ball.image, ball.rect)
        screen.blit(bouncer.image, bouncer.rect)
        pygame.display.flip()

if __name__ == "__main__":
    main()