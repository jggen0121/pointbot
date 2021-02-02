import discord, asyncio, openpyxl, asyncio, random, time #모듈값 불러오기
import os #모듈값 2
from discord.ext import commands

client = discord.Client()

command_prefix="."  # 명령어 접두사


@client.event
async def on_ready():
    print(" ======================================================================= ")
    print("")
    print(client.user.name)
    print("")
    print(" ======================================================================= ")
    print("")
    print(" 봇 실행중..... ")
    print("")
    print(" ======================================================================= ")
    await client.change_presence(activity=discord.Game(".도움말"), status=discord.Status.online)
    print("")
    print(" 상태 표시 완료! ")
    print("")
    print(" ======================================================================= ")
    print("")
    print(" 성공적으로 봇이 시작되었습니다. ")
    print("")
    print(" ======================================================================= ")


@client.event
async def on_message(message):
    if message.content.startswith(f'{command_prefix}자기소개'):
        await message.channel.send("안녕! 나는 얄룽님이 개발해주신! 포인트봇(V3)야! 반가워~ 채팅창에 **phelp** 를 쳐봐!")


    #명령어 도움말 페이지 임베드
    if message.content.startswith(f'{command_prefix}도움'):
        embed = discord.Embed(title="도움말", description="**이 봇은 `LeeSin#5693 - 얄룽` 님에 의해 개발 되었습니다.**", color=0xffffff)
        embed.add_field(name="**도움**", value="**이 메시지를 보여줍니다.**", inline=False)
        embed.add_field(name="**유틸리티**", value="**`개발중`**", inline=False)
        embed.add_field(name="**레벨**", value="**`레벨`, `레벨 (멘션)`, `정보`**", inline=False)
        embed.add_field(name="**포인트**", value="**`포인트`, `포인트 (멘션)`**", inline=False)
        embed.add_field(name="**서버 초대하기**", value="**[클릭!](https://discord.com/invite/D7chQKTxTA)**", inline=False)
        embed.add_field(name="**포인트봇(V3) 초대하기**",value="**[바로가기](https://discord.com/api/oauth2/authorize?client_id=805064278469115905&permissions=8&scope=bot)**", inline=False)
        embed.add_field(name='**참고**', value='**포인트봇(V3) Ver 3.0\n명령어: `.`**', inline=False)
        embed.set_author(name="포인트봇(V3) 도움말",icon_url="https://i.imgur.com/uLDnDU3.png")
        embed.set_thumbnail(url="https://i.imgur.com/uLDnDU3.png")
        await message.channel.send(embed=embed)

    if message.content.startswith(""):
        file = openpyxl.load_workbook("레벨.xlsx")
        sheet = file.active
        xp = [10, 20, 30, 40, 50]  # 레벨(xp)당 포인트(BP)
        i = 1
        while True:
            if sheet["A" + str(i)].value == str(message.author.id):
                sheet["B" + str(i)].value == sheet["B" + str(i)].value + 5
                if sheet["B" + str(i)].value >= xp[sheet["C" + str(i)].value - 1]:
                    sheet["C" + str(i)].value = sheet["C" + str(i)].value + 1
                    await message.channel.send("경험치 UP!\n현재 누적 경험치 : " + str(sheet["C" + str(i)].value) + "\n경험치 : " + str(sheet["B" + str(i)].value))
                file.save("레벨.xlsx")
                break
            if sheet["A" + str(i)].value == None:
                sheet["A" + str(i)].value == str(message.author.id)
                sheet["B" + str(i)].value = 0
                sheet["C" + str(i)].value = 1
                file.save("레벨.xlsx")
                break

            i += 1



access_token = os.environ["BOT_TOKEN"]
client.run(access_token)
