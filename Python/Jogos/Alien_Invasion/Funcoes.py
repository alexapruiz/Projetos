import sys
from time import sleep
import pygame
from bullet import Bullet
from alien import Alien

def check_events(ai_settings, screen, play_button, ship, aliens, bullets):
    """Responde a eventos de pressionamento de teclas e de mouse."""
    for event in pygame.event.get():
        if event.type == pygame.QUIT:
            sys.exit()
        elif event.type == pygame.MOUSEBUTTONDOWN:
            mouse_x, mouse_y = pygame.mouse.get_pos()
            check_play_button(ai_settings, screen, play_button, ship, aliens, bullets, mouse_x, mouse_y)
        elif event.type == pygame.KEYDOWN:
            if event.key == pygame.K_RIGHT:
                # Move a espaçonave para a direita
                if (ship.rect.centerx < 1135):
                    ship.rect.centerx += 1
                ship.moving_right = True
            if event.key == pygame.K_LEFT:
                # Move a espaçonave para a direita
                if (ship.rect.centerx > 66):
                    ship.rect.centerx -= 1
                ship.moving_left = True
            if event.key == pygame.K_UP:
                # Move a espaçonave para a direita
                if (ship.rect.centery > 50):
                    ship.rect.centery -= 1
                ship.moving_up = True
            if event.key == pygame.K_DOWN:
                # Move a espaçonave para a direita
                if (ship.rect.centery < 550):
                    ship.rect.centery += 1
                ship.moving_down = True
            elif event.key == pygame.K_ESCAPE:
                sys.exit()
            elif event.key == pygame.K_SPACE:
                # Cria um novo projétil e o adiciona ao grupo de projéteis
                fire_bullet(ai_settings, screen, ship, bullets)
        elif event.type == pygame.KEYUP:
            if event.key == pygame.K_RIGHT:
                ship.moving_right = False
            if event.key == pygame.K_LEFT:
                ship.moving_left = False
            if event.key == pygame.K_UP:
                ship.moving_up = False
            if event.key == pygame.K_DOWN:
                ship.moving_down = False

def update_screen(ai_settings, screen, ship, aliens, bullets, play_button):
    """Atualiza as imagens na tela e alterna para a nova tela."""
    # Redesenha a tela a cada passagem pelo laço
    screen.fill(ai_settings.bg_color)

    # Redesenha todos os projéteis atrás da espaçonave e dos alienígenas
    for bullet in bullets.sprites():
        bullet.draw_bullet()

    ship.blitme()
    aliens.draw(screen)

    # Deixa a tela mais recente visível
    pygame.display.flip()

def create_fleet(ai_settings, screen, ship, aliens):
    """Cria uma frota completa de alienígenas."""
    # Cria um alienígena e calcula o número de alienígenas em uma linha
    # O espaçamento entre os alienígenas é igual à largura de um alienígena
    alien = Alien(ai_settings, screen)
    alien_width = alien.rect.width
    available_space_x = ai_settings.screen_width - 2 * alien_width
    number_aliens_x = get_number_aliens_x(ai_settings,alien.rect.width)
    number_rows = get_number_rows(ai_settings, ship.rect.height,alien.rect.height)

    # Cria a frota de alienígenas
    for row_number in range(number_rows):
        for alien_number in range(number_aliens_x):
            create_alien(ai_settings, screen, aliens, alien_number,row_number)

def get_number_aliens_x(ai_settings, alien_width):
    """Determina o número de alienígenas que cabem em uma linha."""
    available_space_x = ai_settings.screen_width - 2 * alien_width
    number_aliens_x = int(available_space_x / (2 * alien_width))
    return number_aliens_x

def create_alien(ai_settings, screen, aliens, alien_number, row_number):
    # Cria um alienígena e o posiciona na linha
    alien = Alien(ai_settings, screen)
    alien_width = alien.rect.width
    alien.x = alien_width + 2 * alien_width * alien_number
    alien.rect.x = alien.x
    alien.rect.y = alien.rect.height + 2 * alien.rect.height * row_number
    aliens.add(alien)

def get_number_rows(ai_settings, ship_height, alien_height):
    """Determina o número de linhas com alienígenas que cabem na tela."""
    available_space_y = (ai_settings.screen_height - (3 * alien_height) - ship_height)
    number_rows = int(available_space_y / (2 * alien_height))
    return number_rows

def update_aliens(ai_settings, screen, ship, aliens, bullets):
    """ Verifica se a frota está em uma das bordas e então atualiza as posições de todos os alienígenas da frota."""
    check_fleet_edges(ai_settings, aliens)
    aliens.update()

    # Verifica se houve colisões entre alienígenas e a espaçonave
    if pygame.sprite.spritecollideany(ship, aliens):
        ship_hit(ai_settings, screen, ship, aliens, bullets)

    # Verifica se há algum alienígena que atingiu a parte inferior da tela
    check_aliens_bottom(ai_settings, screen, ship, aliens, bullets)

def check_fleet_edges(ai_settings, aliens):
    """Responde apropriadamente se algum alienígena alcançou uma borda."""
    for alien in aliens.sprites():
        if alien.check_edges():
            change_fleet_direction(ai_settings, aliens)
            break

def change_fleet_direction(ai_settings, aliens):
    """Faz toda a frota descer e muda a sua direção."""
    x=1
    for alien in aliens.sprites():
        alien.rect.y += ai_settings.fleet_drop_speed

    ai_settings.fleet_direction *= -1

def update_bullets(ai_settings, screen, playbutton, ship, aliens, bullets):
    """Atualiza a posição dos projéteis e se livra dos projéteis antigos."""
    bullets.update()
    # Livra-se dos projéteis que desapareceram
    for bullet in bullets.copy():
        if bullet.rect.bottom <= 0:
            bullets.remove(bullet)

    # Verifica se algum projétil atingiu os alienígenas
    # Em caso afirmativo, livra-se do projétil e do alienígena
    collisions = pygame.sprite.groupcollide(bullets, aliens, True, True)

    if len(aliens) == 0:
        #Destrói os projéteis existentes e cria uma nova frota
        bullets.empty()
        create_fleet(ai_settings, screen, ship, aliens)

    check_bullet_alien_collisions(ai_settings, screen, ship, aliens, bullets)

def fire_bullet(ai_settings, screen, ship, bullets):
    """Dispara um projétil se o limite ainda não foi alcançado."""
    # Cria um novo projétil e o adiciona ao grupo de projéteis
    if len(bullets) < ai_settings.bullets_allowed:
        new_bullet = Bullet(ai_settings, screen, ship)
        bullets.add(new_bullet)

def ship_hit(ai_settings, screen, ship, aliens, bullets):
    """Responde ao fato de a espaçonave ter sido atingida por um alienígena."""
    # Atualiza o painel de pontuações
    #sb.prep_ships()

    # Faz uma pausa
    sleep(0.5)

    # Esvazia a lista de alienígenas e de projéteis
    aliens.empty()
    bullets.empty()

    # Cria uma nova frota e centraliza a espaçonave
    create_fleet(ai_settings, screen, ship, aliens)
    ship.center_ship()

    # Faz uma pausa
    sleep(0.5)

def check_aliens_bottom(ai_settings, screen, ship, aliens, bullets):
    """Verifica se algum alienígena alcançou a parte inferior da tela."""
    screen_rect = screen.get_rect()
    for alien in aliens.sprites():
        if alien.rect.bottom >= screen_rect.bottom:
            #Trata esse caso do mesmo modo que é feito quando a espaçonave é atingida
            ship_hit(ai_settings, screen, ship, aliens, bullets)
            break

def check_play_button(ai_settings, screen, play_button, ship, aliens, bullets, mouse_x, mouse_y):
    """Inicia um novo jogo quando o jogador clicar em Play."""
    button_clicked = play_button.rect.collidepoint(mouse_x, mouse_y)
    # Reinicia as configurações do jogo
    ai_settings.initialize_dynamic_settings()

    # Oculta o cursor do mouse
    pygame.mouse.set_visible(False)

    # Esvazia a lista de alienígenas e de projéteis
    aliens.empty()
    bullets.empty()

    # Cria uma nova frota e centraliza a espaçonave
    create_fleet(ai_settings, screen, ship, aliens)
    ship.center_ship()

def check_bullet_alien_collisions(ai_settings, screen, ship, aliens, bullets):
    """Responde a colisões entre projéteis e alienígenas."""
    # Remove qualquer projétil e alienígena que tenham colidido
    collisions = pygame.sprite.groupcollide(bullets, aliens, True, True)
    if collisions:
        check_high_score(sb)
        stats.score += ai_settings.alien_points

    if len(aliens) == 0:
        # Se a frota toda for destruída, inicia um novo nível
        bullets.empty()
        ai_settings.increase_speed()

        # Aumenta o nível
        #stats.level += 1
        #sb.prep_level()

        create_fleet(ai_settings, screen, ship, aliens)

def check_high_score(sb):
    """Verifica se há uma nova pontuação máxima."""
    sb.prep_high_score()