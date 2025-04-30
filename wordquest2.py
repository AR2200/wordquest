import pygame as pg
import sys
import random
import openpyxl as excel
from openpyxl import Workbook

# 初期化
pg.init()
pg.mixer.init()

# 画面サイズ(x,y)
wid = 800
hig = 600
screen = pg.display.set_mode((wid, hig))
WHITE = (255, 255, 255)
BLACK = (0, 0, 0)
GRAY = (100, 100, 100)
GREEN = (0, 255, 0)
FONT = pg.font.Font(None, 36)
pg.display.set_caption("Word Quest")

# テキストボックスの設定
textbox_rect = pg.Rect(250, 285, 300, 50)
input_text = ""
active = False

number_box = pg.Rect(630, 360, 75, 50)

# ボタン
play_btn = pg.image.load("images/play2.png")
play_btn = pg.transform.scale(play_btn, (200, 75))
help_btn = pg.image.load("images/help2.png")
help_btn = pg.transform.scale(help_btn, (200, 75))
menu_btn = pg.image.load("images/menu2.png")
menu_btn = pg.transform.scale(menu_btn, (200, 75))
continue_btn = pg.image.load("images/continue.png")
continue_btn = pg.transform.scale(continue_btn, (200, 75))
select_btn = pg.image.load("images/levelselect.png")
select_btn = pg.transform.scale(select_btn, (200, 75))

level_buttons = {
    # Rect(x, y, width, height)
    "1": pg.Rect(100, 200, 75, 100),
    "2": pg.Rect(230, 200, 75, 100),
    "3": pg.Rect(360, 200, 75, 100),
    "4": pg.Rect(490, 200, 75, 100),
    "Miss": pg.Rect(625, 200, 75, 100)
}

life_buttons = {
    "1": pg.Rect(170, 350, 50, 50),
    "2": pg.Rect(230, 350, 50, 50),
    "3": pg.Rect(290, 350, 50, 50)
}

number_buttons = {
    "10": pg.Rect(500, 350, 50, 50),
    "30": pg.Rect(560, 350, 50, 50),
    "50": pg.Rect(620, 350, 50, 50)
}

# 入力用変数
input_text = ""
input_rect = pg.Rect(wid // 2 - 100, hig // 2 - 25, 200, 50)

# カーソル設定
cursor_visible = True
cursor_counter = 0
cursor_position = 0
CURSOR_INTERVAL = 500  # ミリ秒単位での点滅周期

# 単語データ
workbook = excel.load_workbook("data/wordlist.xlsx")
word = {}
nigate = {}

# デフォルト
c = 0
page = 1
pushFlag = False
life_count = 0
score = 0
correct = 0
incorrect = 0
selected_count = 0
selected_level = None
selected_life = 0
selected_button_level = None
selected_button_life = None
selected_button_count = None
selected_button_practice = None
game_over = False

# ボタン制御
def button_to_jump(btn, newpage):
    global page, pushFlag
    click = pg.mouse.get_pressed()
    (mx, my) = pg.mouse.get_pos()
    if click[0]:
        if btn.collidepoint(mx, my) and pushFlag == False:
            pg.mixer.Sound("sounds/pi.wav").play()
            page = newpage
            pushFlag = True
    else:
        pushFlag = False


# ゲーム画面
def gamestart():
    global cursor_visible
    screen.fill(pg.Color("Black"))
    font = pg.font.Font(None, 100)
    text = font.render("WORD QUEST", True, pg.Color("WHITE"))
    textrect = text.get_rect(center=(wid // 2, 180))
    screen.blit(text, textrect)
    font = pg.font.Font(None, 50)
    screen.blit(text, textrect)
    btn1 = screen.blit(play_btn, (300, 300))
    btn2 = screen.blit(help_btn, (300, 400))

    button_to_jump(btn1, 6)
    button_to_jump(btn2, 2)


def manual():
    pos_list = {
        1: "n: noun(名詞)",
        2: "v: verb(動詞)",
        3: "adj: adjective(形容詞)",
        4: "adv: adverb(副詞)",
        5: "pre: preposition(前置詞)",
        6: "con: conjunction(接続詞)",
    }

    screen.fill(pg.Color("Black"))

    pg.draw.rect(screen, WHITE, textbox_rect, 2)
    font = pg.font.Font("ipaexg.ttf", 25)
    font_surface = font.render("Input: manual", True, WHITE)
    screen.blit(font_surface, (252, 300))

    target_text = font.render("n: 遊び方，マニュアル (m-----)", True, WHITE)
    text_rect = target_text.get_rect(center=(wid // 2, hig // 2 - 100))
    screen.blit(target_text, text_rect)

    word_text = font.render("Words: 1/10 ", True, WHITE)
    screen.blit(word_text, (50, 50))

    score_text = font.render("Score: 10", True, WHITE)
    screen.blit(score_text, (50, 90))

    level_text = font.render("Level: TOEIC 600", True, WHITE)
    screen.blit(level_text, (550, 50))

    life_text = font.render("Life: 3", True, WHITE)
    screen.blit(life_text, (350, 50))

    manual1 = font.render("Parts of Speech →", True, pg.Color("YELLOW"))
    manual2 = font.render("← Japanese Word", True, pg.Color("YELLOW"))
    manual3 = font.render(
        "(Correct : +10, Incorecct : -5, Total: Words + Lifes×10)",
        True,
        pg.Color("YELLOW"),
    )
    manual4 = font.render("← Type your answer.", True, pg.Color("YELLOW"))
    manual5 = font.render("Then push enter key for judge.", True, pg.Color("YELLOW"))
    screen.blit(manual1, (50, hig // 2 - 110))
    screen.blit(manual2, (540, hig // 2 - 110))
    screen.blit(manual3, (50, 130))
    screen.blit(manual4, (535, hig // 2))
    screen.blit(manual5, (425, hig // 2 + 40))

    for i in pos_list.keys():
        pos_text = font.render(pos_list[i], True, pg.Color("YELLOW"))
        screen.blit(pos_text, (50, hig // 2 - 80 + i * 40))

    btn1 = screen.blit(play_btn, (50, 500))
    btn3 = screen.blit(menu_btn, (550, 500))
    button_to_jump(btn1, 6)
    button_to_jump(btn3, 1)

    pg.display.flip()


def create_dict(sheet):
    global word
    word = {}
    try:
        for row in sheet.iter_rows(values_only=True):
            if len(row) >= 2:
                key, value = row[:2]
                word[key] = value
    except Exception as e:
        print("END OF CREATE")
    return word


def levelselect():
    global word, current_word, level_buttons, selected_button_level, selected_button_life, selected_level, selected_button_count, selected_button_practice
    global level, workbook, selected_life, life_count, number_buttons, practice_button, count, number_count
    while True:
        screen.fill(pg.Color("Black"))
        font = pg.font.Font(None, 100)
        level_text = font.render(("LEVEL SELECT"), True, pg.Color("WHITE"))
        textrect = level_text.get_rect(center=(wid // 2, 100))
        screen.blit(level_text, textrect)

        for event in pg.event.get():
            if event.type == pg.QUIT:
                pg.quit()
                sys.exit()
            elif event.type == pg.MOUSEBUTTONDOWN:
                x, y = pg.mouse.get_pos()
                for level, button_rect in level_buttons.items():
                    if button_rect.collidepoint(x, y):
                        selected_button_level = button_rect
                        print(f"{level}が選択されました。")
                        selected_level = str(level)
                        print(str(level))
                        if level == "Miss":
                            word = nigate.copy()
                            if any(nigate) == False:
                                print("There are no words in 'nigate'.")
                                continue

                        else:
                            sheet = workbook[selected_level]
                            word = create_dict(sheet)
                            print(len(word))

                for life, button_rect in life_buttons.items():
                    if button_rect.collidepoint(x, y):
                        selected_button_life = button_rect
                        print(f"lifeは{life}個です。")
                        selected_life = int(life)
                        life_count = selected_life  # (right to left)

                for number, button_rect in number_buttons.items():
                    if button_rect.collidepoint(x, y):
                        selected_button_count = button_rect
                        print(f"問題数は{number}問です。")
                        number_count = int(number)
                        current_word = random.choice(list(word.keys()))
                        count = number_count

        for level, button_rect in level_buttons.items():
            if button_rect == selected_button_level:
                pg.draw.rect(screen, GREEN, button_rect)
            else:
                pg.draw.rect(screen, GRAY, button_rect)
            text_surface = FONT.render(level, True, WHITE)
            text_rect = text_surface.get_rect(center=button_rect.center)
            screen.blit(text_surface, text_rect)

        for life, button_rect in life_buttons.items():
            if button_rect == selected_button_life:
                pg.draw.rect(screen, GREEN, button_rect)
            else:
                pg.draw.rect(screen, GRAY, button_rect)
            text_surface = FONT.render(life, True, WHITE)
            text_rect = text_surface.get_rect(center=button_rect.center)
            screen.blit(text_surface, text_rect)

        for number, button_rect in number_buttons.items():
            if button_rect == selected_button_count:
                pg.draw.rect(screen, GREEN, button_rect)
            else:
                pg.draw.rect(screen, GRAY, button_rect)
            text_surface = FONT.render(number, True, WHITE)
            text_rect = text_surface.get_rect(center=button_rect.center)
            screen.blit(text_surface, text_rect) 
        
        level_text = FONT.render("Level:", True, WHITE)
        screen.blit(level_text, (100, 150))

        life_text = FONT.render("Life:", True, WHITE)
        screen.blit(life_text, (100, 360))

        number_text = FONT.render("Words:", True, WHITE)
        screen.blit(number_text, (400, 360))

        if (
            selected_button_level != None
            and selected_button_life != None
            and selected_button_count != None
            and len(word) != 0
        ):
            btn1 = screen.blit(play_btn, (550, 500))
            button_to_jump(btn1, 3)
        
        btn3 = screen.blit(menu_btn, (50, 500))
        button_to_jump(btn3, 1)

        if page == 1 or (page == 3 and selected_level != None):
            break
        else:
            page == 6

        pg.display.flip()


def gamestage():
    global word, count, correct, incorrect, score, page, game_over, current_word,c 
    global input_text, input_rect, selected_life, life_count, cursor_visible, cursor_counter, cursor_position
    gamereset()
    
    while not game_over:
        screen.fill(pg.Color("Black"))

        # イベント処理
        for event in pg.event.get():
            if event.type == pg.QUIT:
                pg.quit()
                sys.exit()
            elif event.type == pg.KEYDOWN:
                if event.key == pg.K_BACKSPACE:
                    if cursor_position > 0:
                        input_text = input_text[:cursor_position-1] + input_text[cursor_position:]
                        cursor_position -= 1
                elif event.key == pg.K_DELETE:
                    if cursor_position < len(input_text):
                        input_text = input_text[:cursor_position] + input_text[cursor_position+1:]
                elif event.key == pg.K_LEFT:
                    if cursor_position > 0:
                        cursor_position -= 1
                elif event.key == pg.K_RIGHT:
                    if cursor_position < len(input_text):
                        cursor_position += 1

                elif event.key == pg.K_RETURN:
                    c += 1 #正解判定の回数(Enter押された回数)
                    #cursor_visible = False #Enterキーでカーソルを非表示
                    if input_text == word[current_word]:  # correct answer
                        correct += 1
                        score += 10
                        if correct >= count:
                            game_over = True
                            score += selected_life * 10
                            page = 4
                            pg.mixer.Sound("sounds/hatena.wav").play()
                            break
                        else:
                            pg.mixer.Sound("sounds/pon.wav").play()
                            current_word = random.choice(list(word.keys()))
                            
                        input_text = ""
                        cursor_position = 0

                    elif input_text != word[current_word]:  # incorrect answer
                        incorrect += 1
                        score -= 5
                        life_count -= 1

                        if incorrect >= selected_life:
                            game_over = True
                            page = 5
                            pg.mixer.Sound("sounds/down.wav").play()
                            break
                        else:
                            pg.mixer.Sound("sounds/beam.wav").play()

                else:
                    input_text = input_text[:cursor_position] + event.unicode + input_text[cursor_position:]
                    cursor_position += 1


        # 入力テキストの描画
        pg.draw.rect(screen, WHITE, textbox_rect, 2)
        font = pg.font.Font("ipaexg.ttf", 25)
        font_surface = font.render("Input: {}".format(input_text), True, WHITE)
        screen.blit(font_surface, (252, 300))

        target_text = font.render("{} ({})".format(current_word, word[current_word][0] + "-" * (len(word[current_word])-1)), True, WHITE)
        text_rect = target_text.get_rect(center=(wid // 2, hig // 2 - 100))
        screen.blit(target_text, text_rect)

        word_text = font.render("Words: {}/{} ".format(correct, count), True, WHITE)
        screen.blit(word_text, (50, 50))

        score_text = font.render("Score: {}".format(score), True, WHITE)
        screen.blit(score_text, (50, 90))

        level_text = font.render("Level: {}".format(selected_level), True, WHITE)
        screen.blit(level_text, (550, 50))

        life_text = font.render("Life: {}".format(life_count), True, WHITE)
        screen.blit(life_text, (350, 50))

        # カーソルの描画
        if cursor_visible:
            cursor_counter += 1
            if cursor_counter % 30 < 15:  # カーソルの点滅
                cursor_x = wid // 2 - 70 + font.size(input_text[:cursor_position])[0]
                cursor_y = 300
                cursor_height = font_surface.get_height()
                pg.draw.rect(screen, WHITE, (cursor_x, cursor_y, 2, cursor_height))

        pg.display.flip()
            
        # カーソルの点滅速度を調整
        pg.time.Clock().tick(30)


def gameclear():
    screen.fill(pg.Color("Black"))
    font = pg.font.Font(None, 100)
    text = font.render("GAME CLEAR", True, pg.Color("WHITE"))
    textrect = text.get_rect(center=(wid // 2, 100))
    screen.blit(text, textrect)
    font = pg.font.Font(None, 80)
    score_text = font.render(
        "Score: {} ({}/{})".format(score, correct, count), True, WHITE
    )
    textrect = text.get_rect(center=(wid // 2, hig // 2))
    screen.blit(score_text, textrect)
    btn1 = screen.blit(play_btn, (50, 500))
    btn2 = screen.blit(select_btn, (300, 500))
    btn3 = screen.blit(menu_btn, (550, 500))
    button_to_jump(btn1, 3)
    button_to_jump(btn2, 6)
    button_to_jump(btn3, 1)

    pg.display.flip()


def gameover():
    screen.fill(pg.Color("Black"))
    font = pg.font.Font(None, 100)
    text = font.render("GAME OVER", True, pg.Color("WHITE"))
    textrect = text.get_rect(center=(wid // 2, 100))
    screen.blit(text, textrect)
    font = pg.font.Font(None, 80)
    score_text = font.render(
        "Score: {} ({}/{})".format(score, correct, count), True, WHITE
    )
    screen.blit(score_text, (150, 200))
    font = pg.font.Font("ipaexg.ttf", 25)
    text = font.render("{}".format(current_word), True, WHITE)
    screen.blit(text, (150, 275))
    text = font.render("Correct Answer: {}".format(word[current_word]), True, WHITE)
    screen.blit(text, (150, 325))
    text = font.render("Your Answer: {}".format(input_text), True, WHITE)
    screen.blit(text, (150, 375))
    btn1 = screen.blit(continue_btn, (50, 500))
    btn2 = screen.blit(select_btn, (300, 500))
    btn3 = screen.blit(menu_btn, (550, 500))
    button_to_jump(btn1, 3)
    button_to_jump(btn2, 6)
    button_to_jump(btn3, 1)

    pg.display.flip()


def gamereset():
    global pushFlag, count, score, correct, incorrect, game_over, c, current_word, word
    global selected_count, input_text, selected_life, life_count, number_count, cursor_position
    pushFlag = False
    count = number_count
    c = 0
    score = 0
    correct = 0
    incorrect = 0
    life_count = selected_life
    game_over = False
    input_text = ""
    cursor_position = 0
    current_word = random.choice(list(word.keys()))

# メインループ
while True:
    if page == 1:
        gamestart()
    elif page == 2:
        manual()
    elif page == 3:
        gamestage()
    elif page == 4:
        gameclear()
    elif page == 5:
        gameover()
    elif page == 6:
        levelselect()

    # 画面を表示
    pg.display.update()
    pg.time.Clock().tick(60)
    # 閉じるボタンで終了
    for event in pg.event.get():
        if event.type == pg.QUIT:
            pg.quit()
            sys.exit()
