import obspython as obs
import re
import win32gui
import win32api
import threading

# Global variables holding the values of data settings / properties
update_frequency = 50     # Update frequency in msec

# Global animation activity flag
is_enabled = False
scene_name = ""

class ScenesInvalidatedException(Exception):
  pass

scenes_invalidated = False

update_lock = threading.Lock()

def on_event(event):
  global is_enabled, scenes_invalidated, update_lock

  if event == obs.OBS_FRONTEND_EVENT_SCRIPTING_SHUTDOWN:
    update_lock.acquire()
    is_enabled = False
    scenes_invalidated = True
    update_lock.release()
  if (event == obs.OBS_FRONTEND_EVENT_SCENE_COLLECTION_CLEANUP
      or event == obs.OBS_FRONTEND_EVENT_SCENE_CHANGED
      or event == obs.OBS_FRONTEND_EVENT_SCENE_LIST_CHANGED
      or event == obs.OBS_FRONTEND_EVENT_SCENE_COLLECTION_CHANGING
      or event == obs.OBS_FRONTEND_EVENT_SCENE_COLLECTION_CHANGED
      or event == obs.OBS_FRONTEND_EVENT_SCENE_COLLECTION_LIST_CHANGED
      ):
    scenes_invalidated = True

def walk_scene_items_in_current_scene(searched_name, walker):
  source = obs.obs_frontend_get_current_scene()
  if not source:
    return

  scene = obs.obs_scene_from_source(source)
  if not scene:
    return

  walk_scene_items(searched_name, walker, scene)
  obs.obs_source_release(source)

def walk_scene_items(searched_name, walker, scene):
  global scenes_invalidated
  if scenes_invalidated:
    raise ScenesInvalidatedException

  source = obs.obs_scene_get_source(scene)
  if not source:
    return
  source_name = obs.obs_source_get_name(source)
  # print("Walking " + source_name)
  if source_name == searched_name:
    sync_items_in_scene(scene, walker)
    return

  items = obs.obs_scene_enum_items(scene)
  for item in items:
    if scenes_invalidated:
      raise ScenesInvalidatedException
    if obs.obs_sceneitem_is_group(item):
      group_items = obs.obs_sceneitem_group_enum_items(item)
      for grouped_item in group_items:
        if scenes_invalidated:
          raise ScenesInvalidatedException
        grouped_item_as_source = obs.obs_sceneitem_get_source(grouped_item)
        if not grouped_item_as_source:
          continue
        grouped_item_as_scene = obs.obs_scene_from_source(grouped_item_as_source)
        if grouped_item_as_scene:
          walk_scene_items(searched_name, walker, grouped_item_as_scene)
      obs.sceneitem_list_release(group_items)
    else:
      item_as_source = obs.obs_sceneitem_get_source(item)
      if not item_as_source:
        continue
      item_as_scene = obs.obs_scene_from_source(item_as_source)
      if item_as_scene:
        walk_scene_items(searched_name, walker, item_as_scene)
  obs.sceneitem_list_release(items)
  return

def sync_items_in_scene(scene, walker):
  global scenes_invalidated
  # print("syncing items")
  items = obs.obs_scene_enum_items(scene)
  for item in items:
    if scenes_invalidated:
      raise ScenesInvalidatedException
    walker(item, items)
  obs.sceneitem_list_release(items)

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

def sync_scene_item(scene_item, scene_items):
  if not is_window_scene_item(scene_item):
    return

  hwndMain = get_hwnd_by_scene_item(scene_item)
  if not hwndMain:
    return

  is_visible = obs.obs_sceneitem_visible(scene_item)
  if win32gui.IsIconic(hwndMain):
    if is_visible:
      obs.obs_sceneitem_set_visible(scene_item, False)
      return
  else:
    if not is_visible:
      obs.obs_sceneitem_set_visible(scene_item, True)

  if win32gui.GetForegroundWindow() == hwndMain:
    reorder_to_top(hwndMain, scene_item, scene_items)

  sync_scene_item_pos(scene_item, hwndMain)
  return

def is_window_scene_item(scene_item):
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
    return False

  capture_mode = obs.obs_data_get_string(settings, 'capture_mode')
  if capture_mode and capture_mode != "window":
    return False

  return True

def reorder_to_top(hwnd, scene_item, scene_items):
  global scenes_invalidated
  while True:
    if scenes_invalidated:
      raise ScenesInvalidatedException
    any_changed = False
    for other_item in scene_items:
      if obs.obs_sceneitem_get_id(other_item) == obs.obs_sceneitem_get_id(scene_item):
        continue

      otherHwnd = get_hwnd_by_scene_item(other_item)
      if not otherHwnd or otherHwnd == hwnd:
        continue

      if not is_window_scene_item(other_item):
        continue

      while obs.obs_sceneitem_get_order_position(other_item) > obs.obs_sceneitem_get_order_position(scene_item):
        if scenes_invalidated:
          raise ScenesInvalidatedException
        any_changed = True
        obs.obs_sceneitem_set_order(scene_item, 0)
    if not any_changed:
      break

  return

def sync_scene_items():
  global scene_name
  walk_scene_items_in_current_scene(scene_name, sync_scene_item)
  # walk_scene_items_in_scene_by_name(scene_name, sync_scene_item)

lookup = {}
def get_hwnd_by_scene_item(scene_item):
  global lookup
  id = obs.obs_sceneitem_get_id(scene_item)
  if id in lookup:
    if win32gui.IsWindow(lookup[id]):
      return lookup[id]

    del lookup[id]

  hwnd = search_scene_item_hwnd(scene_item)
  if hwnd:
    lookup[id] = hwnd
  return hwnd

def search_scene_item_hwnd(scene_item):
  source = obs.obs_sceneitem_get_source(scene_item)
  if not source:
    return None
  settings = obs.obs_source_get_settings(source)
  if not settings:
    return None
  window = obs.obs_data_get_string(settings, 'window')
  if window != "":
    windowParts = window.split(":")
    windowClassName = unescape_window_name(windowParts[1])
    windowName = unescape_window_name(windowParts[0])
    return win32gui.FindWindow(windowClassName, windowName)

  obs.obs_data_release(settings)

def sync_scene_item_pos(sceneitem, hwnd):
  if not hwnd:
    return

  rect = win32gui.GetWindowRect(hwnd)
  # print(rect)

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

def unescape_window_name(str):
  def repl(m):
    # print(m.group(1))
    ch = chr(int(m.group(1), 16))
    return ch

  return re.sub('#([A-Z0-9]{2})', repl, str)

total_seconds = 0.0
# Called every frame
def script_tick(seconds):
  global total_seconds, update_frequency, is_enabled, scenes_invalidated

  total_seconds += seconds
  if total_seconds < update_frequency / 1000.0:
    return

  update_lock.acquire()
  scenes_invalidated = False

  if not is_enabled:
    update_lock.release()
    return

  total_seconds = 0.0
  
  try:
    sync_scene_items()
  except ScenesInvalidatedException:
    pass

  update_lock.release()

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
  global update_frequency, scene_name, is_enabled
  update_frequency = obs.obs_data_get_double(settings, "update_frequency")
  scene_name = obs.obs_data_get_string(settings, "scene_name")
  is_enabled = obs.obs_data_get_bool(settings, "is_enabled")

def populate_list_property_with_scene_names(list_property):
  scenes = obs.obs_frontend_get_scene_names()
  # print(scenes)
  obs.obs_property_list_clear(list_property)
  obs.obs_property_list_add_string(list_property, "", "")
  for scene in scenes:
    obs.obs_property_list_add_string(list_property, scene, scene)
