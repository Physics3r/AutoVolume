from ctypes import cast, POINTER

from comtypes import CLSCTX_ALL, GUID
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

default_target_volume: int = 80


def modify_volume(target_volume: int):
    # 获取音频设备
    devices = AudioUtilities.GetSpeakers()
    # 获取接口
    # interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    interface = devices.Activate(GUID("{5CDF2C82-841E-4546-9722-0CF74078229A}"), CLSCTX_ALL, None)

    # 转换为接口指针并调节音量
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume_before: int = int(volume.GetMasterVolumeLevelScalar() * 100)
    print("Volume Before: " + str(volume_before))
    print("Target Volume: " + str(default_target_volume))

    volume.SetMasterVolumeLevelScalar(target_volume / 100, None)
    volume_now: int = int(volume.GetMasterVolumeLevelScalar() * 100)
    print("Volume Now: " + str(volume_now))

    # 验证
    if volume_now == target_volume:
        print("Done, and you can close this window now.")
    else:
        print("Sorry, but something wrong happened :)")

    input()


if __name__ == '__main__':
    modify_volume(default_target_volume)
