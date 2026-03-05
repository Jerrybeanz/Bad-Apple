import xlwings as xw
from bad_apple import resource_path, VideoPlayer

if __name__ == "__main__":
    wb = xw.Book()
    wb.save("Bad Apple.xlsx")
    wb.app.activate(steal_focus=True)
    wb.app.screen_updating = False
    sheet = wb.sheets['Sheet1']
    length = 48
    height = 36

    player = VideoPlayer(sheet, length, height)
    player.resize_playback_range(1.4, 12)
    player.format_playback_range()
    player.load_video(resource_path("Bad Apple.mp4"))
    wb.app.screen_updating = True
    player.play_video()