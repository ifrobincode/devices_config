# 简介

devices_config 脚本用于对同类型网络设备批量下发配置使用。

# 使用说明

## Step 1、准备 config_template

- config_template.txt 文件为配置模板，即为需要下发的配置。
- 为每台设备下发的具体配置有所不同（比如设备主机名），这个不相同的配置内容在模板中被定义为变量，例如`{{ hostname }}`。
- 不同的变量，名称一定要各不相同。如果多个配置内容相同可以引用同一变量。

## Step 2、准备 devices_info

- devices_info.xlsx 文件为设备登录信息以及每台设备配置需要的不同变量内容。
- Sheet1 内存放着需要下发配置的网络设备登录信息，比如 IP 地址、Username、Password 等。
- Sheet2 内存放着每台设备对应配置模板所需要的所有变量参数。

## Step 3、执行脚本

1. 首先，脚本会让用户输入设备信息文件（默认：devices_info.xlsx）和配置模板文件（默认：config_template.txt）。
   - 用户可能会有多个设备信息文件和配置模板文件，准备为不同类型的设备下发配置。
2. 脚本通过对两个文件的预查之后会提示用户选择执行模式：
   - Preview（只生成配置文件，不下发配置）。
   - Deploy（生成配置文件并下发配置）。

## 提示

- 输入设备信息文件与配置模板文件时，用户只输入文件名即可，扩展名脚本会自动补全。
- 一个设备的登录信息与需要下发的配置参数，由设备信息文件 Sheet1 与 Sheet2 中的 host 字段关联，请确保两个 Sheet 中 host 字段的一致性。
- 预查（Precheck）过程如果发现问题，会记录在`脚本所在目录/logs/%Y-%m-%d_%H-%M-%S/precheck.log`中。
- 下发配置过程（Deploy）如果发现问题，会记录在`脚本所在目录/logs/%Y-%m-%d_%H-%M-%S/deployment.log`中。

# 更新日志

详见[UPDATE.md](https://github.com/ifrobincode/devices_config/blob/main/UPDATE.md)。