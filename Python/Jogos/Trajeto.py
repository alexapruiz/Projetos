import pygame, random
from pygame.locals import *

# Constantes de direção
UP = 0
RIGHT = 1
DOWN = 2
LEFT = 3

# Helper functions
def DefinePosicaoRandomica():
    x = random.randint(1, 80)
    y = random.randint(1, 59)
    return (x * 10, y * 10)

def Colisao(c1, c2):
    return (c1[0] == c2[0]) and (c1[1] == c2[1])

def DefineDirecao():
    # Procura o caminho mais curto
    if abs(snake[0][0] - apple_pos[0]) < abs(snake[0][1] - apple_pos[1]):
        if (snake[0][0] < apple_pos[0]):
            return RIGHT
        else:
            return LEFT
    else:
        if (snake[0][1] < apple_pos[1]):
            return DOWN
        else:
            return UP

def MoveSerpente():
    # Move a serpente conforme a direção
    if my_direction == UP:
        snake[0] = (snake[0][0], snake[0][1] - 10)
    if my_direction == DOWN:
        snake[0] = (snake[0][0], snake[0][1] + 10)
    if my_direction == RIGHT:
        snake[0] = (snake[0][0] + 10, snake[0][1])
    if my_direction == LEFT:
        snake[0] = (snake[0][0] - 10, snake[0][1])

pygame.init()
screen = pygame.display.set_mode((900, 600))
pygame.display.set_caption('Serpente Automática')

#Cria a serpente e posiciona no centro da tela
snake = [(450, 300), (460, 310), (470, 320)]
snake_skin = pygame.Surface((10, 10))
snake_skin.fill((255, 255, 255))  # White

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
clock = pygame.time.Clock()

while True:
    clock.tick(40)
    for event in pygame.event.get():
        if event.type == QUIT:
            pygame.quit()
            exit()

    # Verifica se a serpente tocou nas bordas
    if (snake[0][0] == 900):
        my_direction = LEFT
    elif (snake[0][0] == 0):
        my_direction = RIGHT
    elif (snake[0][1] == 600):
        my_direction = UP
    elif (snake[0][1] == 0):
        my_direction = DOWN

    if Colisao(snake[0], apple_pos):
        apple_pos = DefinePosicaoRandomica()
        score = score + 1
        my_direction = DefineDirecao()

    for i in range(len(snake) - 1, 0, -1):
        snake[i] = (snake[i - 1][0], snake[i - 1][1])

    MoveSerpente()

    screen.fill((0, 0, 0))
    screen.blit(apple, apple_pos)

    #for x in range(0, 900, 10):  # Draw vertical lines
    #    pygame.draw.line(screen, (40, 40, 40), (x, 0), (x, 900))
    #for y in range(0, 600, 10):  # Draw vertical lines
    #    pygame.draw.line(screen, (40, 40, 40), (0, y), (900, y))

    score_font = font.render('Score: %s' % (score), True, (255, 255, 255))
    score_rect = score_font.get_rect()
    score_rect.topleft = (900 - 100, 10)
    screen.blit(score_font, score_rect)

    for pos in snake:
        screen.blit(snake_skin, pos)

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