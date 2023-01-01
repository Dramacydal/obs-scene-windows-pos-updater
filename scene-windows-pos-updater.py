import obspython as obs
import re
import win32gui
import win32api


# Global variables holding the values of data settings / properties
update_frequency = 50     # Update frequency in msec

# Global animation activity flag
is_enabled = True
scene_name = ""

def walk_scene_items_in_current_source(walker):
  current_scene_as_source = obs.obs_frontend_get_current_scene()
  if not current_scene_as_source:
    return

  current_scene_name = obs.obs_source_get_name(current_scene_as_source)
  if current_scene_name != scene_name:
    return

  current_scene = obs.obs_scene_from_source(current_scene_as_source)
  items = obs.obs_scene_enum_items(current_scene)
  for i, s in enumerate(items):
      walker(s, items)

  obs.sceneitem_list_release(items)
  obs.obs_source_release(current_scene_as_source)
  
def walk_scene_items_in_scene_by_name(walker):
  current_scene_as_source = obs.obs_frontend_get_current_scene()
  if not current_scene_as_source:
    return

  current_scene_name = obs.obs_source_get_name(current_scene_as_source)
  if current_scene_name != scene_name:
    return

  current_scene = obs.obs_scene_from_source(current_scene_as_source)
  items = obs.obs_scene_enum_items(current_scene)
  for i, s in enumerate(items):
      walker(s, items)

  obs.sceneitem_list_release(items)
  obs.obs_source_release(current_scene_as_source)

# Called at script load
def script_load(settings):
  return

def script_save(settings):
  obs.obs_save_sources()

def script_unload():
  return

def script_description():
    return """Scene window position updater.
Synchronizes scene window position with it's real position.
Also toggles item visibilty when window is minimized and reorders currently active window on top.
Updates only top-level window sources.
Any nested scenes or grouped windows will not be updated."""

def sync_scene_item(scene_item, scene_items):
  if not is_window_scene_item(scene_item):
    return

  hwndMain = search_hwnd_by_scene(scene_item)
  if not hwndMain:
    return

  order = obs.obs_sceneitem_get_order_position(scene_item)

  is_visible = obs.obs_sceneitem_visible(scene_item)
  if win32gui.IsIconic(hwndMain):
    if is_visible:
      obs.obs_sceneitem_set_visible(scene_item, False)
      return
  else:
    if not is_visible:
      obs.obs_sceneitem_set_visible(scene_item, True)

  if win32gui.GetForegroundWindow() == hwndMain:
    if (order != len(scene_items)):
      reorder_to_top(scene_item, scene_items)

  # print(hwndMain)
  sync_scene_item_pos(scene_item, hwndMain)
  return

def is_window_scene_item(scene_item):
  source = obs.obs_sceneitem_get_source(scene_item)
  type = obs.obs_source_get_type(source)
  if type != 0:
    return False

  settings = obs.obs_source_get_settings(source)
  # json = obs.obs_data_get_json(settings)
  # print(json)
  window = obs.obs_data_get_string(settings, 'window')
  if not window or window == "":
    return False

  capture_mode = obs.obs_data_get_string(settings, 'capture_mode')
  if capture_mode and capture_mode != "window":
    return False

  return True

def reorder_to_top(scene_item, scene_items):
  while True:
    any_changed = False
    for other_item in scene_items:
      if obs.obs_sceneitem_get_id(other_item) == obs.obs_sceneitem_get_id(scene_item):
        continue

      if not is_window_scene_item(other_item):
        continue

      while obs.obs_sceneitem_get_order_position(other_item) > obs.obs_sceneitem_get_order_position(scene_item):
        any_changed = True
        obs.obs_sceneitem_set_order(scene_item, 0)
    if not any_changed:
      break

  return

def sync_scene_items():
  walk_scene_items_in_current_source(sync_scene_item)

lookup = {}
def search_hwnd_by_scene(scene_item):
  global lookup
  id = obs.obs_sceneitem_get_id(scene_item)
  if id in lookup:
    if win32gui.IsWindow(lookup[id]):
      return lookup[id]

    del lookup[id]

  hwnd = get_scene_item_hwnd(scene_item)
  if hwnd:
    lookup[id] = hwnd
  return hwnd

def get_scene_item_hwnd(scene_item):
  source = obs.obs_sceneitem_get_source(scene_item)
  settings = obs.obs_source_get_settings(source)
  window = obs.obs_data_get_string(settings, 'window')
  if window == "":
    return None
  obs.obs_data_release(settings)

  windowParts = window.split(":")
  windowClassName = unescape_window_name(windowParts[1])
  windowName = unescape_window_name(windowParts[0])

  return win32gui.FindWindow(windowClassName, windowName)

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
  global total_seconds, update_frequency, is_enabled
  total_seconds += seconds
  if total_seconds < update_frequency / 1000.0:
    return

  if not is_enabled:
    return

  total_seconds = 0.0
  sync_scene_items()

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
