# �롼������Excel �� Json���Ѵ����Ƥߤ�

## �Ȥ���

### ����

����1��json�Υ����Ȥʤ�Excel��Range��ɽ��ʸ����
����2���оݤ�Excel����¸����Ƥ���ե����(ʣ����)
����3�����ϸ��json��Ϳ���륭��̾
����4��json �ե�����ν�����

### launch.json

��ư���Υѥ�᡼��

```json
{
    "version": "0.2.0",
    "configurations": [
        {
            "name": ".NET Core Launch (console)",
            "type": "coreclr",
            "request": "launch",
            "preLaunchTask": "build",
            "program": "${workspaceFolder}/bin/Debug/net5.0/ExcelToJson.dll",
            "args": [
                "./config/",
                "./excel/",
                "range",
                "./output"
            ],
            "cwd": "${workspaceFolder}",
            "console": "internalConsole",
            "stopAtEntry": false
        },
        {
            "name": ".NET Core Attach",
            "type": "coreclr",
            "request": "attach"
        }
    ]
}
```

### task.json

��ư���Υ�����

```json
{
    "version": "2.0.0",
    "tasks": [
        {
            "label": "build",
            "command": "dotnet",
            "type": "process",
            "args": [
                "build",
                "${workspaceFolder}/ExcelToJson.csproj",
                "/property:GenerateFullPaths=true",
                "/consoleloggerparameters:NoSummary"
            ],
            "problemMatcher": "$msCompile"
        },
        {
            "label": "publish",
            "command": "dotnet",
            "type": "process",
            "args": [
                "publish",
                "${workspaceFolder}/ExcelToJson.csproj",
                "/property:GenerateFullPaths=true",
                "/consoleloggerparameters:NoSummary"
            ],
            "problemMatcher": "$msCompile"
        },
        {
            "label": "watch",
            "command": "dotnet",
            "type": "process",
            "args": [
                "watch",
                "run",
                "${workspaceFolder}/ExcelToJson.csproj",
                "/property:GenerateFullPaths=true",
                "/consoleloggerparameters:NoSummary"
            ],
            "problemMatcher": "$msCompile"
        }
    ]
}
```

## ���Ϸ��

excel�ե�����˳�Ǽ���줿Excle�ե�������оݤ˼¹�

A1��`Yamada`
A2��`Kento`
A3��`test`

�����Ϥ���Ƥ���Excel�ξ��ϰʲ��Τ褦�˽��Ϥ���롣

```json
{
    "range": [
        {
            "A1": "Yamada"
        },
        {
            "A2": "Kento"
        },
        {
            "A3": "test"
        }
    ]
}
```

