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
parser.read_string(open(db['_config_file']).read())

if 'default_curve' not in db:
    print('[AB Extract] setting default_curve')
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
        self.data_original = list(struct.iter_unpack('<f', bytes.fromhex(curve)))
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
        idx = self.l.index(curve.m[mv])
        for i in range(idx, 0, -1):
            d_idx = self.l[i]['idx']
            self.data[d_idx + 2] = (target_offset,)
        # idx = self.m[mv]['idx']
        # self.data[idx + 2] = (target_offset,)
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
        with open(db['_config_file'], 'w') as cfg:
            parser.write(cfg)
        print('[Apply AB] Restarting AB')
        subprocess.run('taskkill /f /im MSIAfterburner.exe', shell=True)
        subprocess.Popen('"C:\\Program Files (x86)\\MSI Afterburner\\MSIAfterburner.exe"', shell=True)
        time.sleep(2)
        print('[Apply AB] AB restarted')

    def data_apply_optimal(self):
        m = 0
        tmp = {}
        gen_o = db['_final_generation_offset'] if '_final_generation_offset' in db else 0
        for v, o in reversed(sorted(db['_desired_vf_offset'].items(), key=lambda e: float(e[0]))):
            obj = db['vf_offset'][v]
            if float(v) > m:
                m = float(v)
            if 'stable_at' in obj:
                target_offset = o - (OFFSET_STEP * (obj['stable_at'] + gen_o))
                tmp[v] = target_offset
        for v, o in reversed(sorted(db['_desired_vf_offset'].items(), key=lambda e: float(e[0]))):
            obj = db['vf_offset'][v]
            if 'stable_at' in obj:
                target_offset = o - (OFFSET_STEP * (obj['stable_at'] + gen_o))
                # offset at higher voltages should not be more than points at lower voltage
                for tv, to in tmp.items():
                    if float(tv) < float(v) and to < target_offset:
                        target_offset = to
                print(f'[Apply Optimal] Setting {v}mV and below to +{target_offset}MHz (={self.m[v]["mhz"] + target_offset}MHz)')
                idx = self.l.index(self.m[v])
                for i in range(idx, 0, -1):
                    d_idx = self.l[i]['idx']
                    self.data[d_idx + 2] = (target_offset,)
        self.set_max_voltage(str(m))

    def test(self, methods, gpu_is_testing, gpu_is_ended):
        for m in methods:
            print(f'[Test] Testing GPU with {m}')
            if m == 'cp2077':
                def cleanup():
                    try:
                        subprocess.run('taskkill /f /im CrashReporter.exe', shell=True)
                        subprocess.run('taskkill /f /im Cyberpunk2077.exe', shell=True)
                    except:
                        pass

                while True:
                    cleanup()
                    subprocess.Popen('"C:\\Program Files (x86)\\Steam\\steamapps\\common\\Cyberpunk 2077\\bin\\x64\\Cyberpunk2077.exe" -skipStartScreen', shell=True)
                    for i in range(120):
                        try:
                            w = pyautogui.getWindowsWithTitle("Cyberpunk 2077 (C) 2020 by CD Projekt RED")[0]
                            w.activate()
                            print(w)
                            break
                        except:
                            print('[cp2077] Waiting for window')
                        time.sleep(1)
                    else:
                        continue
                    break
                started = False
                gpu_is_testing()
                for i in range(TEST_SECONDS):
                    print(f'[cp2077] {i}/{TEST_SECONDS}')
                    try:
                        pyautogui.getWindowsWithTitle("Cyberpunk 2077 (C) 2020 by CD Projekt RED")[0]
                    except:
                        print('[cp2077] Likely crashed')
                        cleanup()
                        gpu_is_ended()
                        return False
                    pyautogui.press('space')
                    if not started and click_any([f'cp2077/start.png'], 1, 0.8):
                        started = True
                    time.sleep(1)
                print('[cp2077] Seems stable')
                pyautogui.getWindowsWithTitle("Cyberpunk 2077 (C) 2020 by CD Projekt RED")[0].close()
                gpu_is_ended()
            if m.startswith('3dmark'):
                def cleanup():
                    try:
                        subprocess.run('taskkill /f /im 3DMark.exe', shell=True)
                        subprocess.run('taskkill /f /im 3DMarkLauncher.exe', shell=True)
                        subprocess.run('taskkill /f /im SystemInfoHelper.exe', shell=True)
                        subprocess.run('taskkill /f /im javaw.exe', shell=True)
                        subprocess.run('taskkill /f /im FMSISvc.exe', shell=True)
                        time.sleep(5)
                    except Exception as e:
                        print(e)
                    
                while True:
                    # cleanup()
                    subprocess.call("start steam://rungameid/223850", shell=True)
                    for i in range(60):
                        try:
                            pyautogui.getWindowsWithTitle("3DMark Advanced Edition")[0].maximize()
                            break
                        except:
                            print('[3DMark] Waiting for "3DMark Advanced Edition" window')
                        time.sleep(1)
                    print('[3DMark] Clicking "benchmarks"')
                    click_any(['3dmark/b1.png', '3dmark/b2.png'], 5)

                    benchmark = m.split('/')[1]
                    cont = False
                    if benchmark == 'fs':
                        click_any(['3dmark/fs/c.png'], 5)
                        click_any(['3dmark/fs/d.png'], 5)
                        print('[3DMark] Waiting for settings')
                        time.sleep(10)
                        for i in range(1, 6):
                            click_any([f'3dmark/fs/e{i}.png'], 1, 0.99)
                        cont = click_any(['3dmark/fs/f.png'], 5)
                    elif benchmark == 'pr':
                        click_any(['3dmark/pr/c.png'], 5)
                        click_any(['3dmark/pr/d.png'], 5)
                        print('[3DMark] Waiting for settings')
                        time.sleep(10)
                        for i in range(1, 4):
                            click_any([f'3dmark/pr/e{i}.png'], 1, 0.99)
                        cont = click_any(['3dmark/pr/f.png'], 5)
                    if not cont:
                        cleanup()
                        continue
                    for i in range(30):
                        try:
                            w = pyautogui.getWindowsWithTitle("3DMark Workload")[0]
                            print(w)
                            break
                        except:
                            print('[3DMark] Waiting for "3DMark Workload" window')
                            pass
                        time.sleep(1)
                    else:
                        cleanup()
                        continue
                    break
                gpu_is_testing()
                for i in range(TEST_SECONDS):
                    print(f'[3DMark] {i}/{TEST_SECONDS}')
                    try:
                         pyautogui.getWindowsWithTitle("3DMark Workload")[0]
                    except:
                        print('[3DMark] Likely crashed')
                        return False
                    time.sleep(1)
                print('[3DMark] Seems stable')
                pyautogui.getWindowsWithTitle("3DMark Workload")[0].close()
                gpu_is_ended()
                # cleanup()
            if m == 'superposition':
                res = [(1920, 1080), (1280, 720), (2560, 1440)]
                timeout = TEST_SECONDS / round(len(res))
                for w, h in res:
                    cmd = f'"C:\\Program Files\\Unigine\\Superposition Benchmark\\bin\\superposition.exe" -preset 0 -video_app direct3d11 -shaders_quality 3 -textures_quality 2 -dof 1 -motion_blur 1 -video_vsync 0 -video_mode -1 -console_command "world_load superposition/superposition && render_manager_create_textures 1" -project_name Superposition -video_fullscreen 0 -video_width {w} -video_height {h} -extern_plugin GPUMonitor -mode 0 -sound 0 -tooltips 1'
                    args = shlex.split(cmd)
                    start = time.time()
                    try:
                        gpu_is_testing()
                        subprocess.call(cmd, timeout=timeout)
                        gpu_is_ended()
                    except:
                        pass
                    seconds_passed = time.time() - start
                    if seconds_passed < timeout:
                        return False
        return True

    def optimize(self):
        if 'vf_offset' not in db:
            db['vf_offset'] = {}
                    
        for v, o in sorted(db['_desired_vf_offset'].items(), key=lambda e: float(e[0])):
            if v not in db['vf_offset']:
                db['vf_offset'][v] = {}
            obj = db['vf_offset'][v]
            if 'stable_at' in obj:
                target_offset = o - (OFFSET_STEP * obj['stable_at'])
                print(f'[Optimize] {v}mV is already tested and is stable at +{target_offset}MHz')
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
                    print(f'[Optimize] No stable OC at {v}mV')
                    break
                self.data = self.data_original.copy()
                self.set_max_voltage(v)
                self.set_offset(v, target_offset)
                self.display(True)
                time.sleep(2)
                self.apply_to_ab()
                
                def gpu_is_testing():
                    print('[Optimize] GPU test started')
                    obj['is_testing'] = True
                    save_db()

                def gpu_is_ended():
                    print('[Optimiz e] GPU test ended')
                    if 'is_testing' in obj:
                        del obj['is_testing']
                    save_db()

                time.sleep(2)
                is_stable = curve.test(db['_test_methods'], gpu_is_testing, gpu_is_ended)
                gpu_is_ended()
                
                if is_stable:
                    print(f'[Optimize] Assuming +{target_offset}MHz is stable at {v}mV')
                    obj['stable_at'] = obj['current_gen']
                    save_db()
                    break
                obj['current_gen'] += 1
                save_db()
                print(f'[Optimize] +{target_offset}mhz is not stable at {v}mV, moving on')

curve = VFCurve(db['default_curve'])
curve.optimize()
curve.data_apply_optimal()
curve.display(True)
curve.apply_to_ab()
