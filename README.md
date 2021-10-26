# ルールを決めてExcel を Jsonに変換してみる

## 使い方

### 引数

引数1：jsonのキーとなるExcelのRangeを表す文字列
引数2：対象のExcelが保存されているフォルダ(複数可)
引数3：出力後のjsonに与えるキー名
引数4：json ファイルの出力先

### launch.json

起動時のパラメータ

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

起動時のタスク

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

## 出力結果

excelフォルダに格納されたExcleファイルを対象に実行

A1に`Yamada`
A2に`Kento`
A3に`test`

と入力されているExcelの場合は以下のように出力される。

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

