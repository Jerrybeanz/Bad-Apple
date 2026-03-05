import os
import xlwings as xw
from bad_apple import resource_path, VideoPlayer


if __name__ == "__main__":
    base_path = os.path.abspath(".")
    wb_name = "Bad Apple.xlsx"
    wb_path = os.path.join(base_path, wb_name)

    if os.path.isfile(wb_path):
        wb = xw.Book(wb_path)
    else:
        wb = xw.Book()
        wb.save(wb_name)

    wb.app.activate(steal_focus=True)
    wb.app.screen_updating = False
    sheet = wb.sheets.active
    # Video size is 480x360, we use 1/10 for better performance
    length = 48
    height = 36

    player = VideoPlayer(sheet, length, height)
    player.resize_playback_range(1.4, 12)
    player.format_playback_range()
    player.load_video(resource_path("Bad Apple.mp4"))
    wb.app.screen_updating = True
    player.play_video()