import random
import os
import pygame

SCREEN_RECT = pygame.Rect(0, 0, 480, 700)
FRAME_PER_SECOND = 60
IMAGE_PATH = os.path.join(os.getcwd() + '\\飞机大战\\images\\')
CREATE_ENEMY_EVENT = pygame.USEREVENT
FIRE_EVENT = pygame.USEREVENT + 1


class GameSprite(pygame.sprite.Sprite):
    """飞机大战游戏精灵"""

    def __init__(self, image_name, speed=1):

        # 调用父类的初始化方法
        super().__init__()

        # 定义对象的属性
        self.image = pygame.image.load(image_name)
        self.rect = self.image.get_rect()
        self.speed = speed

    def update(self):

        # 在屏幕的垂直方向上移动
        self.rect.y += self.speed


class Background(GameSprite):
    """游戏背景精灵"""

    def __init__(self, is_alt=False):

        # 调用父类方法实现精灵的创建
        bgPath = os.path.join(IMAGE_PATH + 'background.png')
        super().__init__(bgPath)

        if is_alt:
            self.rect.y = -self.rect.height

    def update(self):

        super().update()

        if self.rect.y >= SCREEN_RECT.height:
            self.rect.y = -self.rect.height


class Enemy(GameSprite):
    """敌机精灵"""

    def __init__(self):

        # 1. 调用父类方法，创建敌机精灵，同时指定敌机图片
        enemyPath = os.path.join(IMAGE_PATH + 'enemy1.png')
        super().__init__(enemyPath)

        # 2. 指定敌机的初始随机速度
        self.speed = random.randint(1, 3)

        # 3. 指定敌机的初始随机位置
        self.rect.bottom = 0

        max_x = SCREEN_RECT.width - self.rect.width
        self.rect.x = random.randint(0, max_x)

    def update(self):

        # 1. 调用父类方法，保持垂直方向飞行
        super().update()

        # 2. 判断是否飞出屏幕，如果是，需要从精灵组删除敌机
        if self.rect.y >= SCREEN_RECT.height:
            # print("飞出屏幕...")
            self.kill()

    def __del__(self):
        # print("敌机挂了%s"%self.rect)
        pass


class Hero(GameSprite):
    """英雄精灵"""

    def __init__(self):

        heroPath = os.path.join(IMAGE_PATH + 'me1.png')
        super().__init__(heroPath, 0)

        self.rect.centerx = SCREEN_RECT.centerx
        self.rect.bottom = SCREEN_RECT.bottom - 120

        self.bullets = pygame.sprite.Group()

    def update(self):

        # 英雄在水平方向移动
        self.rect.x += self.speed

        if self.rect.right >= SCREEN_RECT.right:
            self.rect.right = SCREEN_RECT.right
        elif self.rect.x < 0:
            self.rect.x = 0

    def fire(self):
        print("发射子弹")

        for i in range(3):
            # 1. 创建子弹精灵
            bullet = Bullet()

            # 2. 设置精灵位置
            bullet.rect.bottom = self.rect.y - i * 16
            bullet.rect.centerx = self.rect.centerx

            # 3. 将精灵添加到精灵组
            self.bullets.add(bullet)


class Bullet(GameSprite):
    """子弹精灵"""

    def __init__(self):

        bulletPath = os.path.join(IMAGE_PATH + 'bullet1.png')
        super().__init__(bulletPath, -2)

    def update(self):

        # 调用父类方法，让子类沿垂直方向飞行
        super().update()

        # 判断子弹是否飞出屏幕
        if self.rect.bottom < 0:
            self.kill()

    def __del__(self):
        print("子弹销毁...")
