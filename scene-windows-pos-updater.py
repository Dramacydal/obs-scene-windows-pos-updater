from datetime import datetime
import obspython as obs
import re
import win32gui
import win32api
import win32process
import threading
import psutil

class ScenesInvalidatedException(Exception):
  pass

class SceneWindowPosUpdater:
  def __init__(self):
    self.update_frequency = 50     # Update frequency in msec
    self.is_enabled = False
    self.scene_name = ""
    self.scenes_invalidated = False
    self.update_lock = threading.Lock()
    self.lookup = {}

  def walk_scene_items_in_current_scene(self, searched_name, walker):
    source = obs.obs_frontend_get_current_scene()
    if not source:
      return

    scene = obs.obs_scene_from_source(source)
    if not scene:
      return

    self.walk_scene_items(searched_name, walker, scene)
    obs.obs_source_release(source)

  def walk_scene_items(self, searched_name, walker, scene):
    if self.scenes_invalidated:
      raise ScenesInvalidatedException

    source = obs.obs_scene_get_source(scene)
    if not source:
      return
    source_name = obs.obs_source_get_name(source)
    # print("Walking " + source_name)
    if source_name == searched_name:
      self.sync_items_in_scene(scene, walker)
      return

    items = obs.obs_scene_enum_items(scene)
    for item in items:
      if self.scenes_invalidated:
        raise ScenesInvalidatedException
      if obs.obs_sceneitem_is_group(item):
        group_items = obs.obs_sceneitem_group_enum_items(item)
        for grouped_item in group_items:
          if self.scenes_invalidated:
            raise ScenesInvalidatedException
          grouped_item_as_source = obs.obs_sceneitem_get_source(grouped_item)
          if not grouped_item_as_source:
            continue
          grouped_item_as_scene = obs.obs_scene_from_source(grouped_item_as_source)
          if grouped_item_as_scene:
            self.walk_scene_items(searched_name, walker, grouped_item_as_scene)
        obs.sceneitem_list_release(group_items)
      else:
        item_as_source = obs.obs_sceneitem_get_source(item)
        if not item_as_source:
          continue
        item_as_scene = obs.obs_scene_from_source(item_as_source)
        if item_as_scene:
          self.walk_scene_items(searched_name, walker, item_as_scene)
    obs.sceneitem_list_release(items)
    return

  def sync_items_in_scene(self, scene, walker):
    self.log("syncing items")
    items = obs.obs_scene_enum_items(scene)
    for item in items:
      if self.scenes_invalidated:
        raise ScenesInvalidatedException
      walker(item, items)
    obs.sceneitem_list_release(items)

  def sync_scene_item(self, scene_item, scene_items):
    if not self.is_window_scene_item(scene_item):
      return

    source = obs.obs_sceneitem_get_source(scene_item)
    if source == None:
      return
    
    hwndMain = self.get_hwnd_by_scene_item(scene_item)
    if not hwndMain:
      return

    self.log("Syncing HWND " + str(hwndMain))
    is_visible = obs.obs_sceneitem_visible(scene_item)
    self.log("Is visible: " + str(is_visible))
    if win32gui.IsIconic(hwndMain):
      if is_visible:
        obs.obs_sceneitem_set_visible(scene_item, False)
        return
    else:
      if not is_visible:
        obs.obs_sceneitem_set_visible(scene_item, True)

    self.log("Is foreground: " + str(win32gui.GetForegroundWindow() == hwndMain))

    if win32gui.GetForegroundWindow() == hwndMain:
      self.reorder_to_top(hwndMain, scene_item, scene_items)

    self.sync_scene_item_pos(scene_item, hwndMain)
    return

  def get_scene_item_uniq_id(self, scene_item):
    source = obs.obs_sceneitem_get_source(scene_item)
    if not source:
      return ""
    type = obs.obs_source_get_type(source)
    if type != 0:
      return ""

    settings = obs.obs_source_get_settings(source)
    if not settings:
      return ""
    window = obs.obs_data_get_string(settings, 'window')
    if not window or window == "":
      obs.obs_data_release(settings)
      return ""
    
    return window

  def is_window_scene_item(self, scene_item):
    source = obs.obs_sceneitem_get_source(scene_item)
    if not source:
      return False
    type = obs.obs_source_get_type(source)
    if type != 0:
      return False

    settings = obs.obs_source_get_settings(source)
    if not settings:
      return False
    window = obs.obs_data_get_string(settings, 'window')
    if not window or window == "":
      obs.obs_data_release(settings)
      return False

    capture_mode = obs.obs_data_get_string(settings, 'capture_mode')
    if capture_mode and capture_mode != "window":
      obs.obs_data_release(settings)
      return False

    obs.obs_data_release(settings)
    return True

  def is_scene_item_captured(self, scene_item):
    source = obs.obs_sceneitem_get_source(scene_item)
    if source == None:
      return False
    
    return obs.obs_source_get_height(source) != 0

  def reorder_to_top(self, hwnd, scene_item, scene_items):
    while True:
      if self.scenes_invalidated:
        raise ScenesInvalidatedException
      any_changed = False
      for other_item in scene_items:
        if obs.obs_sceneitem_get_id(other_item) == obs.obs_sceneitem_get_id(scene_item):
          continue

        if not self.is_window_scene_item(other_item):
          continue

        otherHwnd = self.get_hwnd_by_scene_item(other_item)
        if not otherHwnd or otherHwnd == hwnd:
          continue

        while obs.obs_sceneitem_get_order_position(other_item) > obs.obs_sceneitem_get_order_position(scene_item):
          if self.scenes_invalidated:
            raise ScenesInvalidatedException
          any_changed = True
          obs.obs_sceneitem_set_order(scene_item, 0)
      if not any_changed:
        break

    return

  def sync_scene_items(self):
    self.walk_scene_items_in_current_scene(self.scene_name, self.sync_scene_item)
    # walk_scene_items_in_scene_by_name(scene_name, sync_scene_item)

  def get_hwnd_by_scene_item(self, scene_item):
    if not self.is_scene_item_captured(scene_item):
      return None
    
    uniq_id = self.get_scene_item_uniq_id(scene_item)

    id = obs.obs_sceneitem_get_id(scene_item)
    self.log("get_hwnd_by_scene_item: " + str(id))
    if uniq_id in self.lookup:
      self.log("Cached HWND for id " + str(id) + ": " + str(self.lookup[uniq_id]))
      if win32gui.IsWindow(self.lookup[uniq_id]):
        return self.lookup[uniq_id]

      del self.lookup[uniq_id]

    hwnd = self.search_scene_item_hwnd(scene_item)
    if hwnd:
      self.lookup[uniq_id] = hwnd
    return hwnd

  def search_scene_item_hwnd(self, scene_item):
    source = obs.obs_sceneitem_get_source(scene_item)
    if not source:
      return None
    settings = obs.obs_source_get_settings(source)
    if not settings:
      return None

    hwnd = None

    window = obs.obs_data_get_string(settings, 'window')
    self.log("Searching HWND by window: '" + window + "'")
    if window != "":
      windowParts = window.split(":")
      windowClassName = self.unescape_window_name(windowParts[1])
      windowName = self.unescape_window_name(windowParts[0])
      self.log("Window name: '" + windowName + "', classname: '" + windowClassName + "'")
      if windowName != "":
        hwnd = win32gui.FindWindow(windowClassName, windowName)
        self.log("By window name hwnd: " + str(hwnd))

      if hwnd == None or hwnd == 0:
        processName = windowParts[len(windowParts) - 1]
        self.log("Searching process '" + processName + "', classname: '" + windowClassName + "'")
        hwnd = self.get_process_hwnd(processName, windowClassName)     

    obs.obs_data_release(settings)
    return hwnd

  def get_pid(self, processName):
    for p in psutil.process_iter():
      try:
          if p.name() == processName:
            return p.pid
      except (psutil.AccessDenied, psutil.ZombieProcess):
          pass
      except psutil.NoSuchProcess:
          pass
    return None
  
  def get_process_hwnd(self, processName, className):
    pid = self.get_pid(processName)
    self.log("Pid: " + str(pid))
    if pid == None:
      return None
    
    mainHwnd = self.find_window_for_pid(pid, className)
    self.log("hwnd: " + str(mainHwnd))
    return mainHwnd
  
  def find_window_for_pid(self, pid, className):
    result = None
    def callback(hwnd, _):
        try:
          nonlocal result
          ctid, cpid = win32process.GetWindowThreadProcessId(hwnd)
          if cpid == pid and win32gui.GetClassName(hwnd) == className:
              result = hwnd
              return True
          return True
        except:
          return True
    win32gui.EnumWindows(callback, None)
    return result

  def sync_scene_item_pos(self, sceneitem, hwnd):
    if not hwnd:
      return

    rect = win32gui.GetWindowRect(hwnd)

    hMonitor = win32api.MonitorFromWindow(hwnd, 0)
    if not hMonitor:
      return

    monitorInfo = win32api.GetMonitorInfo(hMonitor)
    if not monitorInfo:
      return

    newPos = obs.vec2()
    oldPos = obs.vec2()
    newPos.x = rect[0] - monitorInfo['Monitor'][0]
    newPos.y = rect[1] - monitorInfo['Monitor'][1]

    obs.obs_sceneitem_get_pos(sceneitem, oldPos)
    if oldPos.x != newPos.x or oldPos.y != newPos.y:
      obs.obs_sceneitem_set_pos(sceneitem, newPos)

  def unescape_window_name(self, str):
    def repl(m):
      ch = chr(int(m.group(1), 16))
      return ch

    return re.sub('#([A-Z0-9]{2})', repl, str)
  
  def log(self, msg):
    return
    print("[" + datetime.today().strftime('%Y-%m-%d %H:%M:%S') + "] " + msg)


updater = SceneWindowPosUpdater()

# Global variables holding the values of data settings / properties
update_frequency = 50     # Update frequency in msec

# Global animation activity flag
is_enabled = False
scene_name = ""

scenes_invalidated = False

update_lock = threading.Lock()

def on_event(event):
  global updater

  if event == obs.OBS_FRONTEND_EVENT_SCRIPTING_SHUTDOWN:
    updater.update_lock.acquire()
    updater.is_enabled = False
    updater.scenes_invalidated = True
    updater.update_lock.release()
  if (event == obs.OBS_FRONTEND_EVENT_SCENE_COLLECTION_CLEANUP
      or event == obs.OBS_FRONTEND_EVENT_SCENE_CHANGED
      or event == obs.OBS_FRONTEND_EVENT_SCENE_LIST_CHANGED
      or event == obs.OBS_FRONTEND_EVENT_SCENE_COLLECTION_CHANGING
      or event == obs.OBS_FRONTEND_EVENT_SCENE_COLLECTION_CHANGED
      or event == obs.OBS_FRONTEND_EVENT_SCENE_COLLECTION_LIST_CHANGED
      ):
    updater.scenes_invalidated = True

# Called at script load
def script_load(settings):
  obs.obs_frontend_add_event_callback(on_event)
  return

def script_save(settings):
  obs.obs_save_sources()

def script_unload():
  return

def script_description():
    return """Scene window position updater.
Synchronizes scene window position with it's real position.
Also toggles item visibilty when window is minimized and reorders currently active window on top.
If currently active scene is not the one selected in script settings, scripts walks all scene sources recursively and searches for selected scene.
"""

total_seconds = 0.0
# Called every frame
def script_tick(seconds):
  global total_seconds, updater

  total_seconds += seconds
  if total_seconds < updater.update_frequency / 1000.0:
    return

  updater.update_lock.acquire()
  updater.scenes_invalidated = False

  if not updater.is_enabled:
    updater.update_lock.release()
    return

  total_seconds = 0.0
  
  try:
    updater.sync_scene_items()
  except ScenesInvalidatedException:
    pass

  updater.update_lock.release()

# Called to set default values of data settings
def script_defaults(settings):
  obs.obs_data_set_default_string(settings, "scene_name", "")
  obs.obs_data_set_default_double(settings, "update_frequency", 50.0)
  obs.obs_data_set_default_bool(settings, "is_enabled", True)

# Called to display the properties GUI
def script_properties():
  props = obs.obs_properties_create()
  obs.obs_properties_add_bool(props, "is_enabled", "Enabled")
  obs.obs_properties_add_float_slider(props, "update_frequency", "Update frequency, ms", 10, 1000, 10)

  list_property = obs.obs_properties_add_list(props, "scene_name", "Scene name",
              obs.OBS_COMBO_TYPE_LIST, obs.OBS_COMBO_FORMAT_STRING)
  populate_list_property_with_scene_names(list_property)

  obs.obs_properties_add_button(props, "button", "Refresh list of scenes",
    lambda props,prop: True if populate_list_property_with_scene_names(list_property) else True)

  return props

# Called after change of settings including once after script load
def script_update(settings):
  updater.update_frequency = obs.obs_data_get_double(settings, "update_frequency")
  updater.scene_name = obs.obs_data_get_string(settings, "scene_name")
  updater.is_enabled = obs.obs_data_get_bool(settings, "is_enabled")

def populate_list_property_with_scene_names(list_property):
  scenes = obs.obs_frontend_get_scene_names()
  # print(scenes)
  obs.obs_property_list_clear(list_property)
  obs.obs_property_list_add_string(list_property, "", "")
  for scene in scenes:
    obs.obs_property_list_add_string(list_property, scene, scene)
