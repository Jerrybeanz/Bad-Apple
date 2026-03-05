import xlwings as xw
from bad_apple import resource_path, VideoPlayer


if __name__ == "__main__":
    with xw.App(visible=True) as app:
        app.activate(steal_focus=True)
        app.screen_updating = False
        wb = app.books.active
        sheet = wb.sheets.active
        # Video size is 480x360, we use 1/10 for better performance
        length = 48
        height = 36

        player = VideoPlayer(sheet, length, height)
        player.resize_playback_range(1.4, 12)
        player.format_playback_range()
        player.load_video(resource_path("Bad Apple.mp4"))
        app.screen_updating = True
        player.play_video()
        print("Playback ended...")