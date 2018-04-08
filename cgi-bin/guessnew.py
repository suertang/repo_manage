def main():
    int_guess=checkinput(input("我有一个数字，你猜是几，请输入数字（按8可以退出）: "))
    while(int_guess!=8):
        int_guess=checkinput(input("你输入的数字{1}比我想的{0}哦. 再猜猜看:".format("大" if int_guess>8 else "小",int_guess)))
    print("噢哟，不错哦.猜对了，其实我早就告诉你答案了。")

def checkinput(data):
    try:
        return int(data)
    except:
        print("你没有遵守游戏规则！不理你了！")
        exit(-1)

if __name__=="__main__":
    print(__file__)
    main()