import pygame, random
from pygame.locals import *

# Constantes de direção
UP = 0
RIGHT = 1
DOWN = 2
LEFT = 3
Direcao = 4
x = 0

# Helper functions
def DefinePosicaoRandomica():
    x = random.randint(1, 80)
    y = random.randint(1, 59)
    return x * 10, y * 10


def Colisao(c1, c2):
    return (c1[0] == c2[0]) and (c1[1] == c2[1])


def DefineDirecao():
    # Procura o caminho mais curto
    if abs(snake[0][0] - apple_pos[0]) < abs(snake[0][1] - apple_pos[1]):
        if snake[0][0] < apple_pos[0]:
            return RIGHT
        else:
            return LEFT
    else:
        if snake[0][1] < apple_pos[1]:
            return DOWN
        else:
            return UP


def MoveSerpente(x):
    # Move a serpente conforme a direção
    if (x % 2) == 0:
        if my_direction == UP:
            snake[0] = (snake[0][0], snake[0][1] - 10)
        if my_direction == DOWN:
            snake[0] = (snake[0][0], snake[0][1] + 10)
        if my_direction == RIGHT:
            snake[0] = (snake[0][0] + 10, snake[0][1])
        if my_direction == LEFT:
            snake[0] = (snake[0][0] - 10, snake[0][1])


def MoveJogador(direcao):
    # Move o Jogador
    if direcao == UP:
        player[0] = (player[0][0], player[0][1] - 10)
    if direcao == DOWN:
        player[0] = (player[0][0], player[0][1] + 10)
    if direcao == RIGHT:
        player[0] = (player[0][0] + 10, player[0][1])
    if direcao == LEFT:
        player[0] = (player[0][0] - 10, player[0][1])


def VerificaBordas(objeto):
    if (objeto[0][0] == 900):
        return LEFT
    elif (objeto[0][0] < 0):
        return RIGHT
    elif (objeto[0][1] == 600):
        return UP
    elif (objeto[0][1] < 0):
        return DOWN
    else:
        return 999


pygame.init()
screen = pygame.display.set_mode((900, 600))
pygame.display.set_caption('Serpente Automática')

#Cria a serpente e posiciona no centro da tela
snake = [(450, 300), (460, 310), (470, 320)]
snake_skin = pygame.Surface((10, 10))
snake_skin.fill((255, 255, 0))  # Amarelo

#Cria o jogador e posiciona do lado esquerdo da tela
player = [(100, 100)]
player_skin = pygame.Surface((10, 10))
player_skin.fill((255, 255, 255))  # White

#Define a posição do alvo
apple_pos = DefinePosicaoRandomica()
apple = pygame.Surface((10, 10))
apple.fill((255, 0, 0))

#Define a direção de acordo com a posição do alvo
if (int(snake[0][0]) > int(apple_pos[0])):
    my_direction = LEFT
else:
    my_direction = RIGHT

font = pygame.font.Font('freesansbold.ttf', 18)
score = 0
score_player = 0
clock = pygame.time.Clock()
gameover = False

while (gameover == False):
    clock.tick(10)
    for event in pygame.event.get():
        if event.type == QUIT:
            pygame.quit()
            exit()

        if event.type == KEYDOWN:
            if event.key == K_UP:
                Direcao = UP
            if event.key == K_DOWN:
                Direcao = DOWN
            if event.key == K_LEFT:
                Direcao = LEFT
            if event.key == K_RIGHT:
                Direcao = RIGHT
            if event.key == K_ESCAPE:
                pygame.quit()

    # Verifica se a serpente tocou nas bordas
    retorno = VerificaBordas(snake)
    if (retorno != 999):
        my_direction = retorno

    # Verifica se o jogador tocou nas bordas
    if (VerificaBordas(player) != 999):
        gameover = True

    if Colisao(snake[0], apple_pos):
        apple_pos = DefinePosicaoRandomica()
        score = score + 1
        my_direction = DefineDirecao()

    if Colisao(player[0], apple_pos):
        apple_pos = DefinePosicaoRandomica()
        score_player = score_player + 1

    if Colisao(player[0], snake[0]):
        gameover = True
        break

    for i in range(len(snake) - 1, 0, -1):
        snake[i] = (snake[i - 1][0], snake[i - 1][1])

    MoveSerpente(x)

    x = x + 1
    MoveJogador(Direcao)

    screen.fill((0, 0, 0))
    screen.blit(apple, apple_pos)

    #Atualiza o placar do Computador
    score_font = font.render('Computador: %s' % (score), True, (255, 255, 255))
    score_rect = score_font.get_rect()
    score_rect.topleft = (750, 10)
    screen.blit(score_font, score_rect)

    # Atualiza o placar do Jogador
    score_font = font.render('Alex: %s' % (score_player), True, (255, 255, 255))
    score_rect = score_font.get_rect()
    score_rect.topleft = (10, 10)
    screen.blit(score_font, score_rect)

    #Desenha a serpente
    for pos in snake:
        screen.blit(snake_skin, pos)

    #Desenha o jogador
    for pos in player:
        screen.blit(player_skin, pos)

    #Verifica se precisa atualizar a direção
    if (int(snake[0][0]) == int(apple_pos[0])):
        #Encontrou a linha, agora precisa saber a coluna
        if (int(snake[0][1]) > int(apple_pos[1])):
            my_direction = UP
        else:
            my_direction = DOWN
    else:
        if (int(snake[0][1]) == int(apple_pos[1])):
            #Encontrou a coluna, agora precisa saber a linha
            if (int(snake[0][0]) > int(apple_pos[0])):
                my_direction = LEFT
            else:
                my_direction = RIGHT

    pygame.display.update()

while True:
    game_over_font = pygame.font.Font('freesansbold.ttf', 75)
    game_over_screen = game_over_font.render('Game Over', True, (255, 255, 255))
    game_over_rect = game_over_screen.get_rect()
    game_over_rect.midtop = (900 / 2, 10)
    screen.blit(game_over_screen, game_over_rect)
    pygame.display.update()
    pygame.time.wait(500)
    while True:
        for event in pygame.event.get():
            if event.type == QUIT:
                pygame.quit()
                exit()
