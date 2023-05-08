# obs-scene-windows-pos-updater

Synchronizes scene window position with it's real position on screen.

Also toggles item visibilty when window is minimized and reorders currently active window on top.

If currently active scene is not the one selected in script settings, scripts walks all scene sources recursively and searches for selected scene.

Requires Python 3 (x64 for x64 obs, x32 for x32 obs) and pywin32, install the latest with
- pip3 install pywin32

[![Demo](https://img.youtube.com/vi/1ejjVGxSwW0/0.jpg)](https://www.youtube.com/watch?v=1ejjVGxSwW0)
