#!/usr/bin/env python3
import struct
import json
import subprocess
import configparser
import os
import time
import shlex
import pyautogui

TEST_SECONDS = 3600
OFFSET_STEP = 15

db = {}

def save_db():
    with open('db.json', 'w') as f:
        f.write(json.dumps(db, indent=4, sort_keys=True))
        os.fsync(f)

db = json.loads(open('db.json').read())
save_db()

parser = configparser.ConfigParser()
parser.read_string(open(db['config_file']).read())
if 'default_curve' not in db:
    print('[Main] setting default_curve')
    db['default_curve'] = parser['Startup']['VFCurve']
    save_db()

def click_any(images, timeout, confidence=0.9):
    start = time.time()
    while True:
        if time.time() - start > timeout:
            print(f'[Click] Timeout clicking {images}')
            return False
        try:
            for im in images:
                x = pyautogui.locateCenterOnScreen(im, confidence=confidence)
                print('[Click]', im, x)
                if x is None:
                    continue
                pyautogui.click(x)
                return True
        except:
            pass
    
class VFCurve:
    def __init__(self, curve):
        self.data = list(struct.iter_unpack('<f', bytes.fromhex(curve)))
        self.parse_data()

    def parse_data(self):
        l = []
        m = {}
        for i, v in enumerate(self.data):
            if i <= 2:
                continue
            if v[0] == 0 and self.data[i+1][0] == 0:
                break
            if v[0] < 680:
                continue
            if i % 3 == 0:
                m[str(v[0])] = {'idx': i, 'v': v[0], 'mhz': self.data[i+1][0], 'offset': self.data[i+2][0]}
                l.append(m[str(v[0])])
        self.l = l
        self.m = m       

    def encode(self):
        packed = struct.pack(f'<{len(self.data)}f', *map(lambda d: d[0], self.data))
        return packed.hex()

    def set_offset(self, mv, target_offset):
        idx = self.m[mv]['idx']
        self.data[idx + 2] = (target_offset,)
        self.parse_data()

    def set_max_voltage(self, mv):
        idx = self.l.index(self.m[mv])
        for i in range(idx + 1, len(self.l)):
            d_idx = self.l[i]['idx']
            self.data[d_idx + 2] = (-1000,)
        self.parse_data()

    def display(self, require_offset=False):
        for x in self.l:
            if require_offset and x['offset'] <= 0:
                continue
            print(x)

    def apply_to_ab(self):
        parser['Startup']['coreclkboost'] = '0'
        parser['Startup']['VFCurve'] = self.encode()
        with open(db['config_file'], 'w') as cfg:
            parser.write(cfg)
        print('[VF] Restarting AB')
        subprocess.run('taskkill /f /im MSIAfterburner.exe', shell=True)
        subprocess.Popen('"C:\\Program Files (x86)\\MSI Afterburner\\MSIAfterburner.exe"', shell=True)
        time.sleep(2)
        print('[VF] AB restarted')

    def be_optimal(self):
        m = 0
        for v, o in reversed(sorted(db['desired_vf_offset'].items(), key=lambda e: float(e[0]))):
            obj = db['vf_offset'][v]
            if float(v) > m:
                m = float(v)
            if 'stable_at' in obj:
                target_offset = o - (OFFSET_STEP * obj['stable_at'])
                print(f'[Main] Setting {v}mV and below to +{target_offset}MHz')
                idx = self.l.index(curve.m[v])
                for i in range(idx, 0, -1):
                    d_idx = self.l[i]['idx']
                    self.data[d_idx + 2] = (target_offset,)
        self.set_max_voltage(str(m))
                        
                    

    def test(self, fn):
        print('[VF] Starting test')
        while True:
            subprocess.call("start steam://rungameid/223850", shell=True)
            for i in range(60):
                try:
                    pyautogui.getWindowsWithTitle("3DMark Advanced Edition")[0].maximize()
                    break
                except:
                    print('[VF] Waiting for "3DMark Advanced Edition" window')
                    pass
                time.sleep(1)
            print('[VF] Clicking "benchmarks"')
            click_any(['images/b1.png', 'images/b2.png'], 5)

            benchmark = 'pr'
            if benchmark == 'fs':
                click_any(['images/fs/c.png'], 5)
                click_any(['images/fs/d.png'], 5)
                print('[VF] Waiting for settings')
                time.sleep(10)
                for i in range(1, 6):
                    click_any([f'images/fs/e{i}.png'], 1, 0.99)
                click_any(['images/fs/f.png'], 5)
            elif benchmark == 'pr':
                click_any(['images/pr/c.png'], 5)
                click_any(['images/pr/d.png'], 5)
                print('[VF] Waiting for settings')
                time.sleep(10)
                for i in range(1, 4):
                    click_any([f'images/pr/e{i}.png'], 1, 0.99)
                click_any(['images/pr/f.png'], 5)
            for i in range(30):
                try:
                    w = pyautogui.getWindowsWithTitle("3DMark Workload")[0]
                    print(w)
                    break
                except:
                    print('[VF] Waiting for "3DMark Workload" window')
                    pass
                time.sleep(1)
            else:
                continue
            break
        fn()
        for i in range(TEST_SECONDS):
            try:
                 pyautogui.getWindowsWithTitle("3DMark Workload")[0]
            except:
                print('[VF] Likely crashed')
                return False
            time.sleep(1)
        pyautogui.getWindowsWithTitle("3DMark Workload")[0].close()
        return True
        """
        for w, h in [(1920, 1080), (1280, 720), (2560, 1440)]:
            cmd = f'"C:\\Program Files\\Unigine\\Superposition Benchmark\\bin\\superposition.exe" -preset 0 -video_app direct3d11 -shaders_quality 3 -textures_quality 2 -dof 1 -motion_blur 1 -video_vsync 0 -video_mode -1 -console_command "world_load superposition/superposition && render_manager_create_textures 1" -project_name Superposition -video_fullscreen 0 -video_width {w} -video_height {h} -extern_plugin GPUMonitor -mode 0 -sound 0 -tooltips 1'
            args = shlex.split(cmd)
            start = time.time()
            try:
                subprocess.call(cmd, timeout=TEST_SECONDS)
            except:
                pass
            seconds_passed = time.time() - start
            if seconds_passed < TEST_SECONDS:
                return False
        return True
        """
        

# curve = VFCurve(db['default_curve'])
# curve.display()

print("[Main] Starting in 5 seconds")
# time.sleep(5)

if 'vf_offset' not in db:
    db['vf_offset'] = {}
            
for v, o in sorted(db['desired_vf_offset'].items(), key=lambda e: float(e[0])):
    if v not in db['vf_offset']:
        db['vf_offset'][v] = {}
    obj = db['vf_offset'][v]
    if 'stable_at' in obj:
        target_offset = o - (OFFSET_STEP * obj['stable_at'])
        print(f'[Main] {v}mV is already tested and is stable at +{target_offset}MHz')
        continue
    while True:
        if 'current_gen' not in obj:
            obj['current_gen'] = 0
        if 'is_testing' in obj and obj['is_testing']:
            # system rebooted/failed previously
            del obj['is_testing']
            obj['current_gen'] += 1
        save_db()

        target_offset = o - (OFFSET_STEP * obj['current_gen'])
        if target_offset < 0:
            print(f'[Main] No stable OC at {v}mV')
            break
        curve = VFCurve(db['default_curve'])
        curve.set_max_voltage(v)
        curve.set_offset(v, target_offset)
        curve.display(True)
        curve.apply_to_ab()
        
        def actual_running():
            print('[Main] Real test started')
            obj['is_testing'] = True
            save_db()
            
        is_stable = curve.test(actual_running)

        if 'is_testing' in obj:
            del obj['is_testing']
        if is_stable:
            print(f'[Main] Assuming +{target_offset}MHz is stable at {v}mV')
            obj['stable_at'] = obj['current_gen']
            save_db()
            break
        obj['current_gen'] += 1
        save_db()
        print(f'[Main] +{target_offset}mhz is not stable at {v}mV, moving on')
        time.sleep(5)

print(f'[Main] Setting optimal curve')
curve = VFCurve(db['default_curve'])
curve.be_optimal()
curve.parse_data()
# curve.display(True)
curve.apply_to_ab()
