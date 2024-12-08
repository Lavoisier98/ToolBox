import pyautogui as pg
import pyperclip as pc
import time as tm
import random as rd
import math as mt

def imgSearching(imageName, *args, **kwargs):
    #Optionals: conf, heights, widths, gScale
    conf = kwargs.get("conf",0.85)
    heights = kwargs.get("heights",0.5)
    widths = kwargs.get("widths",0.5)
    gScale = kwargs.get("gScale",True)
    confTest = kwargs.get("confTest",False)
    img = None
    if confTest:
        conf = 0.99
    while img == None:
        try:
            img = pg.locateOnScreen(folder+imageName,grayscale=gScale,confidence=conf)
        except pg.ImageNotFoundException:
            img = None
            if not confTest:
                tm.sleep(0.3)
            else:
                tm.sleep(0.1)
                conf -= 0.01
    confMax = conf
    if heights != 0.5 or widths != 0.5:
        pg.moveTo(x=img.left + widths * img.width, y=img.top + heights * img.height,duration=1.5,tween=pg.easeInQuad)

    if confTest:
        img2 = None
        while img2 == None:
            try:
                img2 = pg.locateOnScreen(folder+imageName,grayscale=gScale,confidence=conf)
            except pg.ImageNotFoundException:
                img2 = None
                tm.sleep(0.1)
                conf -= 0.01
            if not (abs(img.left - img2.left) > 15 or abs(img.top - img2.top) > 15):
                img2 = None
                tm.sleep(0.1)
                conf -= 0.01
        if heights == 0.5 and widths == 0.5:
            pg.moveTo(img2)
    confMin = conf
    return {
        "top": img.top + heights * img.height,
        "left": img.left + widths * img.width,
        "height": img.height,
        "width": img.width,
        "img": img,
        "confMax": confMax,
        "confMin": confMin
    }

def painter(**kwargs):

    colorChangeLimiter = rd.randint(1,75)
    while True:

        #Storaging original position:
        iPoint = pg.position()
        moveY = rd.randint(int(0.301 * tela.height), int(0.847 * tela.height))
        moveX = rd.randint(int(0.054 * tela.width), int(0.98 * tela.width))
        fPoint = pg.position(x=moveX, y=moveY)
        moveY = rd.randint(int(0.301 * tela.height), int(0.847 * tela.height))
        moveX = rd.randint(int(0.054 * tela.width), int(0.98 * tela.width))
        Point2 = pg.position(x=moveX, y=moveY)
        moveY = rd.randint(int(0.301 * tela.height), int(0.847 * tela.height))
        moveX = rd.randint(int(0.054 * tela.width), int(0.98 * tela.width))
        Point3 = pg.position(x=moveX, y=moveY)

        #definir se vai mudar de cor
        colorChange = rd.randint(10,100)
        if colorChange >= colorChangeLimiter:
            colorChangeLimiter = rd.randint(1,75)
            pg.press(["alt","e","c"])
            colorChange = str(rd.randint(0,255))
            pc.copy(colorChange)
            pg.press(["tab","tab","tab","tab","tab","tab","del"])
            pg.hotkey("ctrl","v")
            pg.press(["tab","del"])
            colorChange = str(rd.randint(0,255))
            pc.copy(colorChange)
            pg.hotkey("ctrl","v")
            pg.press(["tab","del"])
            colorChange = str(rd.randint(0,255))
            pc.copy(colorChange)
            pg.hotkey("ctrl","v")
            pg.press(["tab","tab","tab","tab","enter"])
        
        #traçar pontos da curva de dois pontos
        pg.dragTo(x=fPoint.x,y=fPoint.y,duration=2)
        pg.click(Point2.x,Point2.y)
        pg.click(Point3.x,Point3.y)

        #fazer clicar fora do quadradão para desselecionar a curva
        while True:
            moveY = rd.randint(int(0.301 * tela.height), int(0.847 * tela.height))
            moveX = rd.randint(int(0.054 * tela.width), int(0.98 * tela.width))
            if (moveY > max([iPoint.y,fPoint.y,Point2.y,Point3.y]) + 15 or \
                moveY < min([iPoint.y,fPoint.y,Point2.y,Point3.y]) - 15) or \
                (moveX > max([iPoint.x,fPoint.x,Point2.x,Point3.x]) + 15 or \
                moveX < min([iPoint.x,fPoint.x,Point2.x,Point3.x]) - 15):
                break
        pg.click(x=moveX,y=moveY)

        #fazer voltar ao ponto final da curva
        pg.moveTo(x=fPoint.x, y=fPoint.y)
        tm.sleep(1)

def coordinates(**kwargs):
    local = pg.position()
    img = kwargs.get("img",None)
    confTest = kwargs.get("confTest",False)
    locationTest = kwargs.get("position",False)
    intHeights = kwargs.get("heights",0.5)
    intwidths = kwargs.get("widths",0.5)
    if img == None:
        if locationTest:
            pg.moveTo()
        return {
            "xPercent": round(local.x / tela.width, 3),
            "yPercent": round(local.y / tela.height, 3)
        }
    elif not confTest:
        search = imgSearching(img,heights=intHeights,widths=intwidths)
        pg.moveTo(x=search["left"],y=search["top"])
        return {
            "widths": round((local.x - search["img"].left) / search["img"].width, 3),
            "heights": round((local.y - search["img"].top) / search["img"].height, 3),
            "formula": "heights= "+str(round((local.y - search["img"].top) / search["img"].height, 3)) \
                        +", widths= "+str(round((local.x - search["img"].left) / search["img"].width, 3))
        }
    else:
        search = imgSearching(img, confTest=True,heights=intHeights,widths=intwidths)
        pg.moveTo(x=search["left"],y=search["top"])
        return {
            "widths": round((local.x - search["img"].left) / search["img"].width, 3),
            "heights": round((local.y - search["img"].top) / search["img"].height, 3),
            "confMax": round(search["confMax"] - 0.04, 3),
            "confMin": round(search["confMin"] + 0.04, 3),
            "formula": "heights= "+str(round((local.y - search["img"].top) / search["img"].height, 3)) \
                        +", widths= "+str(round((local.x - search["img"].left) / search["img"].width, 3))
        }

def onlymoving():
    while True:
        randomEvent = rd.randint(1, 100)
        if randomEvent > 70:
            pg.press("win")
            tm.sleep(0.5)
            pg.press("win")
        pg.moveTo(x=rd.randint(int(0.054 * tela.width), int(0.98 * tela.width)), y=rd.randint(int(0.2 * tela.height), int(0.8 * tela.height)), duration=2)
        tm.sleep(1)

def timeDefiner(**Inicial):
    #Function para definir timeLimit: Se vazia, retorna o segundo do ano em Integer;
    #Se preenchida, retorna a diferença (em segundos) entre o tempo inicial do kwarg e o tempo final
    timeI = Inicial.get("Inicial", 0)
    currentTime = tm.localtime()[5] + tm.localtime()[4] * 60 + tm.localtime()[3] * 3600 + tm.localtime()[7] * 86400
    return currentTime - timeI
folder = r'C:\Users\NE20119\OneDrive - The Dow Chemical Company\Desktop\Base de Dados\VS Code\Python\Images'
folder2 = r'C:\Users\NE20119\OneDrive - The Dow Chemical Company\Desktop\Base de Dados\VS Code\Python\Screenshots'

print(tm.localtime())
#tI = timeDefiner()
#tm.sleep(4.2)
#timeDiff = timeDefiner(Inicial=tI)
#print(timeDiff)
tm.sleep(4.5)
#pg.scroll(120) #Scroll Test
#print(pc.paste())
tela = pg.size()
#print(float(pc.paste().strip().replace(",",".")))
#print("1;2,5.4/7 ".replace("/",";").replace(".",";").replace(",",";").split(";"))
#print(coordinates(img=r"\SimulatorLoaded.png", confTest=False)) #Location & Conf Tests ,heights= 0.571, widths= 2.124
#print(coordinates()) #X and Y Coordinates in screen percentual
#im2 = pg.screenshot(folder2+'\my_screenshot.png') #Screenshot Test
#painter() #Always Moving with clicks in paint screen
onlymoving() #Always Moving just moving mouse and pressing Win



pass
#teste:
#Ver se dá alguma das notificações restantes
#imgMsg = imgSearching( \
#    imageName=["\VA01-6A.png","\VA01-6B.png","\VA01-6C.png","\VA01-6D.png"], \
#    action="find")
#verify = False
#MaxMMM = ""
#while not verify:
#    imgMsg = imgSearching( \
#        imageName=["\VA01-6B.png","\VA01-6A.png","\VA01-6C.png","\VA01-6D.png"], action="try")
#    if imgMsg["name"] == None:
#        verify = True
#    else:
        #Correção se foi para A e tinha que ir para B, pois imagem A também tem B
#        if imgMsg["name"] in ["\VA01-6A.png"]:
#            if MaxMMM == "":
#                tm.sleep(1.1)
#                print("Passou aqui")
#                imgMsg2 = imgSearching(imageName=["\VA01-6B.png"],action="try")
#                if imgMsg2["name"] != None:
#                    imgMsg = imgMsg2
#        if imgMsg["name"] == "\VA01-6B.png":
#            #MMM Checking
#            tm.sleep(0.9)
#            MaxMMM = mouseCopy(["\VA01-6B.png"], 0.882, 2.327, "click").strip()
#            pg.press("enter")
#            print(MaxMMM)
#            tm.sleep(1.4)
#        elif imgMsg["name"] == "\VA01-6C.png":
#            #Availability Checking
#            tm.sleep(0.3)
#            availabilityCheck = imgSearching(imageName=["\VA01-10.png"],action="try",conf=0.95)
#            if availabilityCheck["img"] != None:
#                parameters["Availability"]["value"] = "Missing"
#            else:
#                #PENDENTE FAZER TESTE DE AVAILABILITY COM CENÁRIO QUE TENHA AVAILABILITY
#                parameters["Availability"]["value"] = "Coming Soon"
#            pg.hotkey("shift","fn","f6")
#            if MaxMMM == "":
#                tm.sleep(0.9)
#            else:
#                tm.sleep(0.3)
#        else:
#            pg.press("enter")
#            if MaxMMM == "":
#                tm.sleep(1.1)
#            else:
#                tm.sleep(0.5)







