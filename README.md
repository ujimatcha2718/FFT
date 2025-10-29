本スクリプトは，Excel入力ファイルを読み込み，**時間波形のFFT解析**および**周波数伝達関数 H(f) = F_out / F_in** を求めるツールである．  

---

# FFT解析スクリプト 説明書

**ファイル名:** `fft.py`  
**作成日:** 2025-10-29  
**対象:** ExcelデータからFFT解析と周波数伝達関数を算出するPythonスクリプト  

---

## 🧭 概要

本スクリプトは，Excel入力ファイルを読み込み，**時間波形のFFT解析**および**周波数伝達関数 H(f) = F_out / F_in** を求めるツールである．  
主な処理の流れは以下のとおりである．

1. Excelファイルの読み込み  
2. サンプル数の調整（切り詰め／ゼロパディング）  
3. 窓関数（Bohman）適用  
4. FFTの実行と正規化・補正  
5. 伝達関数の算出  
6. 各種グラフの生成と保存  

---

## ⚙️ 依存関係

以下のPythonライブラリが必要である．

```bash
pip install numpy scipy pandas matplotlib openpyxl
````

（環境によっては `xlrd` が必要になる場合もある．）

---

## 📂 入力形式

* 入力は **Excelファイル（.xls / .xlsx）**
* ヘッダ行付きの表形式を想定
* データフレームの **2列目を出力チャネル，3列目を入力チャネル** として扱う
* 欠損値は自動で `0` で補完される
* 入力信号にはスケーリング処理 `×0.001` が適用される（例: mV → V）

---

## 🚀 主な機能

1. 時間波形・周波数スペクトル・伝達関数のプロットを自動生成
2. 帯域通過フィルタを適用し，逆FFTで再構成した信号をExcelに出力
3. すべての結果を指定フォルダ（既定: `fft_outputs`）に保存

---

## 🧰 コマンドライン引数

| オプション          | 説明                       | 既定値             |
| -------------- | ------------------------ | --------------- |
| `-i, --input`  | 入力Excelファイル名             | `A_yes_1.9.xls` |
| `-o, --outdir` | 出力ディレクトリ                 | `fft_outputs`   |
| `--N`          | サンプル数                    | `500`           |
| `--dt`         | サンプリング間隔 [s]             | `0.005`         |
| `--fc`         | 帯域上限 [Hz]                | `2.2`           |
| `--fh`         | 帯域下限 [Hz]                | `1.7`           |
| `--time_xlim`  | 時間表示範囲 (min max)         | `0 2.5`         |
| `--freq_xlim`  | 周波数表示範囲 (min max)        | `0 5`           |
| `--no-show`    | グラフの表示を抑制（保存のみ）          | -               |
| `--save-tf`    | 伝達関数の図を保存                | -               |
| `--prefix`     | 出力ファイル名の接頭辞              | -               |
| `--img-format` | 画像出力形式 (`pdf` または `png`) | `pdf`           |
| `--log-level`  | ログ出力レベル                  | `INFO`          |

---

## 💡 使用例

### 1. デフォルト設定で実行

```bash
python fft.py
```

### 2. サンプル数・時間間隔を変更して実行

```bash
python fft.py -i mydata.xlsx --N 1024 --dt 0.001 -o results --img-format png
```

### 3. 確認プロンプト，プロット表示を省略して実行

```bash
python fft.py --no-prompt --no-show
```

---

## 📈 入出力の注意点

* 入力ファイルが指定されない場合，自動で最初の `.xls` / `.xlsx` を探索する．
  意図しないファイルが選ばれる可能性があるため，**明示指定を推奨**する．
* 入力データの列位置が異なる場合は，スクリプト内で列指定を変更する必要がある．
* 窓関数は **Bohman** を採用しており，SciPyのバージョンによる差異がある場合がある．
* FFT結果は**窓補正・正規化済み**だが，解析目的に応じて見直すことを推奨する．

---

## 🗂️ 出力ファイル

* **再構成信号（逆FFT結果）**：Excelファイルで保存
* **プロット画像**：`pdf` または `png` 形式で出力ディレクトリに保存
* **ログ出力**：標準出力に進行状況を表示

---

## ⚠️ 注意事項（信頼性に関して）

* 本スクリプトは**教育・実験用途**向けであり，高精度解析には未対応の部分がある．
* FFTの正規化・窓損失補正・位相処理などに関して，厳密な誤差補償を行っていない．
* 入力信号に自動スケーリング（×0.001）を行うため，**単位系の確認が必須**である．
* 各種ライブラリ（NumPy, SciPy, pandas, matplotlib）のバージョンにより挙動が異なる場合がある．
* **先輩のコードをChatGPTに魔改造させているので動作がおかしい可能性あり**

---

## 🔧 改良提案

* 入出力列を引数で指定可能にする
* Welch法などによる平均スペクトル推定の導入
* 窓関数の種類や正規化方法を選択可能にする
* 振幅・位相の分離出力機能を追加する

---

## 🧾 ライセンスと著作権

* 本READMEは説明文書であり，スクリプト本体`fft.py`の著作権は作成者に帰属する．
* 再配布・改変時は著作者の許諾条件に従うこと．

---

## 📚 参考文献

* 本READMEはスクリプト本体 [`fft.py`](fft.py) を基に作成した．

---
```
"""
fft.py

簡易 FFT / 周波数解析スクリプト

処理の流れ:
 1. 入力ファイル（Excel）を読み込む（1行目はヘッダとしてスキップ）
 2. 指定サンプル数 N にあわせてデータ長を切り詰め/ゼロパディング
 3. 窓関数（Bohman）を適用
 4. FFT を取り、正規化や窓の補正を行う
 5. 入力/出力から周波数伝達関数 H(f)=F_out/F_in を計算
 6. 時間波形 / 周波数スペクトル / 伝達関数（振幅・位相）をプロット

依存: numpy, scipy, pandas, matplotlib, openpyxl/xlrd
"""

import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from scipy import signal
import os
import glob
import argparse
import sys
import logging

# --------- デフォルト設定（コマンドライン引数で上書き可能） ----------
DEFAULT_INPUT_FILE = "A_yes_1.9.xls"  # デフォルトの入力ファイル名（.xls/.xlsx を想定）
DEFAULT_N = 500               # サンプル数
DEFAULT_DT = 0.005            # サンプリング間隔 [s]
DEFAULT_FC = 2.2              # カットオフ周波数（高い方）[Hz]
DEFAULT_FH = 1.7              # カットオフ周波数（低い方）[Hz]
DEFAULT_TIME_XLIM = (0, 2.5)  # 時間表示範囲 [s]
DEFAULT_FREQ_XLIM = (0, 5)    # 周波数表示範囲 [Hz]
DEFAULT_OUTPUT_DIR = "fft_outputs"  # 出力ファイルをまとめるフォルダ


def to_n_length(arr, N):
    """配列 arr を長さ N に合わせる。
    - 短ければゼロでパディング
    - 長ければ先頭 N 要素を切り出す
    """
    a = np.asarray(arr, dtype=float)
    if a.size < N:
        return np.pad(a, (0, N - a.size), 'constant')
    return a[:N]


def read_input_file(input_file: str):
    """入力ファイルを読み込み、(DataFrame, resolved_input_path) を返す。
    - Excel (.xls/.xlsx) は pandas.read_excel
    - ファイルが存在しない場合はカレントディレクトリから候補を自動選択
    """
    # 自動検出: .xls/.xlsx のみを候補として探す（CSV はサポート外とする）
    if not os.path.exists(input_file):
        # まずはプロジェクト内のよく使うサブフォルダを探索（あれば採用）
        alt_dir = os.path.join(os.getcwd(), "2025_B班_後期")
        if os.path.isdir(alt_dir):
            candidates = []
            for ext in ("*.xls", "*.xlsx"):
                candidates.extend(glob.glob(os.path.join(alt_dir, ext)))
            if candidates:
                input_file = candidates[0]
                print(f"Note: configured input not found, using detected file in '{alt_dir}': {input_file}")

    if not os.path.exists(input_file):
        candidates = glob.glob("*.xls") + glob.glob("*.xlsx")
        if candidates:
            input_file = candidates[0]
            print(f"Note: configured input not found, using detected file: {input_file}")
        else:
            # 明示的に .xls/.xlsx を要求する
            raise FileNotFoundError(f"No .xls/.xlsx input file found (tried configured '{input_file}'). Please provide an Excel (.xls/.xlsx) file.")

    # 読み込み本体 (.xls/.xlsx を前提)
    if not input_file.lower().endswith(('.xls', '.xlsx')):
        raise ValueError(f"Input file must be .xls or .xlsx: {input_file}")
    df = pd.read_excel(input_file, header=0)

    return df, input_file


def parse_args(argv=None):
    """コマンドライン引数をパースして返す。argv を渡せばテスト可能。"""
    p = argparse.ArgumentParser(description="Simple FFT + transfer function analysis")
    p.add_argument("-i", "--input", default=DEFAULT_INPUT_FILE, help="入力ファイル (.xls/.xlsx).")
    p.add_argument("-o", "--outdir", default=DEFAULT_OUTPUT_DIR, help="出力ディレクトリ")
    p.add_argument("--N", type=int, default=DEFAULT_N, help="サンプル数 N")
    p.add_argument("--dt", type=float, default=DEFAULT_DT, help="サンプリング間隔 [s]")
    p.add_argument("--fc", type=float, default=DEFAULT_FC, help="帯域上限 [Hz]")
    p.add_argument("--fh", type=float, default=DEFAULT_FH, help="帯域下限 [Hz]")
    p.add_argument("--time_xlim", type=float, nargs=2, default=DEFAULT_TIME_XLIM, help="時間表示範囲 (min max)")
    p.add_argument("--freq_xlim", type=float, nargs=2, default=DEFAULT_FREQ_XLIM, help="周波数表示範囲 (min max)")
    p.add_argument("--no-show", dest="show", action="store_false", help="プロット表示を抑制してファイルのみ保存する")
    p.add_argument("--save-tf", dest="save_tf", action="store_true", help="伝達関数のプロットをファイルに保存する")
    p.add_argument("--prefix", default="", help="出力ファイル名の接頭辞を指定（例: 20251029_）")
    p.add_argument("--log-level", default="INFO", choices=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"], help="ログの詳細レベル")
    p.add_argument("--img-format", default="pdf", choices=["png", "pdf"], help="出力画像フォーマット（png または pdf）")
    p.add_argument("--no-prompt", dest="prompt", action="store_false", help="デフォルト入力の確認プロンプトを表示しない（デフォルトは確認する）")
    return p.parse_args(argv)


def main(argv=None):
    """メイン処理: データ読み込み → FFT → プロット"""
    args = parse_args(argv)

    # パラメータをローカル変数に展開
    input_file = args.input
    output_dir = args.outdir
    N = args.N
    DT = args.dt
    FC = args.fc
    FH = args.fh
    TIME_XLIM = tuple(args.time_xlim)
    FREQ_XLIM = tuple(args.freq_xlim)
    SHOW_PLOTS = args.show
    SAVE_TF = args.save_tf
    PREFIX = args.prefix
    LOG_LEVEL = args.log_level
    IMG_FORMAT = args.img_format
    PROMPT_DEFAULT = args.prompt

    # ロギング設定
    numeric_level = getattr(logging, LOG_LEVEL.upper(), logging.INFO)
    logging.basicConfig(level=numeric_level, format='%(levelname)s: %(message)s')

    # 時間 / 周波数軸の作成
    t = np.arange(0, N * DT, DT)
    freq = np.linspace(0, 1.0 / DT, N)

    # 窓関数（Bohman）: SciPy のバージョン差に対応するフォールバック
    try:
        window = signal.bohman(N)
    except Exception:
        try:
            window = signal.windows.bohman(N)
        except Exception:
            window = signal.get_window('bohman', N)

    # --- 明示的に入力ファイルを指定した場合のみ確認プロンプトを出す（対話端末でのみ）。
    #      デフォルトのまま使用する場合は尋ねません。--no-prompt でプロンプトを完全に無効化できます。 ---
    if input_file != DEFAULT_INPUT_FILE and PROMPT_DEFAULT and sys.stdin.isatty():
        while True:
            resp = input(f"Use input file '{input_file}'? [Y/n] ").strip()
            if resp == "" or resp.lower() in ("y", "yes"):
                logging.info("Using input file: %s", input_file)
                break
            if resp.lower() in ("n", "no"):
                new = input("Enter alternative input file path (or press Enter to cancel): ").strip()
                if new == "":
                    logging.error("No input file selected; aborting.")
                    return
                input_file = new
                logging.info("User selected input file: %s", input_file)
                break
            print("Please answer 'y' or 'n'.")

    # 入力ファイルを読み込む（存在しなければ自動検出）
    try:
        df, resolved_input = read_input_file(input_file)
    except FileNotFoundError as e:
        logging.error("%s", e)
        return
    except Exception as e:
        logging.error("Could not read input file: %s", e)
        return

    # データ列: 期待する列は "列1=出力(インデックス1), 列2=入力(インデックス2)"（0 ベースで言うと位置 1,2）
    # 欠損は 0 埋め、入力は 0.001 倍（mV→V 等の単位スケーリングを既存ロジックに合わせる）
    f = to_n_length(df.iloc[:, 1].fillna(0).astype(float).values, N)
    fIn = to_n_length(df.iloc[:, 2].fillna(0).astype(float).values * 0.001, N)

    # 窓掛け
    f = f * window
    fIn = fIn * window

    # 簡易情報表示
    logging.info("FIn_max: %s [mV]", str(np.amax(np.abs(fIn)) * 1000))

    # FFT
    F = np.fft.fft(f)
    FIn = np.fft.fft(fIn)

    # 正規化: 周波数成分のスケーリングを既存コードのロジックに合わせる
    F = F / (N / 2)
    F[0] = F[0] / 2
    F = F * (N / np.sum(np.abs(window)))

    F_temp = np.abs(F)[0:int(N/2)]
    logging.info("freq_max: %s [Hz]", str(freq[np.argmax(F_temp)]))

    FIn = FIn / (N / 2)
    FIn[0] = FIn[0] / 2
    FIn = FIn * (N / np.sum(np.abs(window)))

    logging.debug("max |F| = %s", np.max(np.abs(F)))
    logging.debug("max |FIn| = %s", np.max(np.abs(FIn)))

    # 周波数伝達関数 H(f) = F_out / F_in（ゼロ除算回避）
    eps = 1e-12
    H = np.zeros_like(F, dtype=complex)
    nonzero = np.abs(FIn) > eps
    H[nonzero] = F[nonzero] / FIn[nonzero]

    # 正の周波数側のみを扱う
    freq_half = freq[0:int(N/2)]
    H_half = H[0:int(N/2)]

    # フィルター（帯域選択）: F2/F2In を作成して逆 FFT で時間信号を得る
    F2 = F.copy()
    F2In = FIn.copy()
    F2[(freq > FC)] = 0
    F2[(freq < FH)] = 0
    F2In[(freq > FC)] = 0
    F2In[(freq < FH)] = 0

    f2 = np.real(np.fft.ifft(F2) * N)
    f2In = np.real(np.fft.ifft(F2In) * N)

    # --- 時間信号の縦軸レンジを処理後の信号に合わせる ---
    try:
        combined = np.concatenate([f2.ravel(), f2In.ravel()])
        y_min = float(np.min(combined))
        y_max = float(np.max(combined))
        if y_max == y_min:
            y_min -= 1.0
            y_max += 1.0
        else:
            margin = 0.05 * (y_max - y_min)
            y_min -= margin
            y_max += margin
    except Exception:
        y_min, y_max = None, None

    # 出力ディレクトリを作成して、出力ファイルをそこに保存する
    try:
        os.makedirs(output_dir, exist_ok=True)
    except Exception as e:
        logging.warning("Could not create output directory '%s': %s", output_dir, e)

    try:
        base = os.path.splitext(os.path.basename(resolved_input))[0]
        out_xlsx = os.path.join(output_dir, f"{PREFIX}{base}_fft.xlsx")
        pd.DataFrame(f2In).to_excel(out_xlsx, sheet_name="bn2.4", index=False, header=False)
        # 追記（もしシートを追加したければ）
        with pd.ExcelWriter(out_xlsx, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
            pd.DataFrame(f2In).to_excel(writer, sheet_name="bah_in", index=False, header=False)
        logging.info("Wrote output Excel: %s", out_xlsx)
    except Exception as e:
        logging.warning("Could not write Excel file %s: %s", out_xlsx, e)

    # ----------------- プロット: 処理前（時間+周波数） -----------------
    plt.figure(figsize=(10, 6))
    plt.rcParams["font.family"] = "Times New Roman"
    plt.rcParams["font.size"] = 12
    plt.subplots_adjust(hspace=0.35)

    # 上段: 時間信号（元）
    plt.subplot(211)
    plt.plot(t, f, label="$V_{ch1}$ [mV]")
    plt.plot(t, fIn, label="$V_{ch2}$ [V]")
    plt.ylabel("Voltage")
    plt.grid(True)
    plt.xlim(*TIME_XLIM)
    if y_min is not None and y_max is not None:
        plt.ylim(y_min, y_max)
    plt.legend(loc="upper right", fontsize=9)
    # 下段: 周波数（元スペクトル）
    plt.subplot(212)
    plt.plot(freq, np.abs(F), label="|F_ch1| (original)")
    plt.plot(freq, np.abs(FIn), label="|F_ch2| (original)")
    plt.xlabel("Frequency [Hz]")
    plt.ylabel("Amplitude")
    plt.grid(True)
    plt.xlim(*FREQ_XLIM)
    plt.legend(loc="upper right", fontsize=9)
    plt.tight_layout()
    # 保存（処理前プロット）
    try:
        pre_img = os.path.join(output_dir, f"{PREFIX}{base}_before.{IMG_FORMAT}")
        plt.savefig(pre_img, dpi=200)
        logging.info("Saved plot: %s", pre_img)
    except Exception as e:
        logging.warning("Could not save plot %s: %s", pre_img, e)

    # ----------------- プロット: 処理後（時間+周波数） -----------------
    plt.figure(figsize=(10, 6))
    plt.rcParams["font.family"] = "Times New Roman"
    plt.rcParams["font.size"] = 12
    plt.subplots_adjust(hspace=0.35)

    # 上段: 時間信号（処理後）
    plt.subplot(211)
    plt.plot(t, f2, label="$V_{ch1}$ processed [mV]")
    plt.plot(t, f2In, label="$V_{ch2}$ processed [V]")
    plt.xlabel("Time [s]")
    plt.ylabel("Voltage")
    plt.grid(True)
    plt.xlim(*TIME_XLIM)
    if y_min is not None and y_max is not None:
        plt.ylim(y_min, y_max)
    plt.legend(loc="upper right", fontsize=9)

    # 下段: 周波数（処理後スペクトル）
    plt.subplot(212)
    plt.plot(freq, np.abs(F2), label="|F_ch1| (processed)")
    plt.plot(freq, np.abs(F2In), label="|F_ch2| (processed)")
    plt.xlabel("Frequency [Hz]")
    plt.ylabel("Amplitude")
    plt.grid(True)
    plt.xlim(*FREQ_XLIM)
    plt.legend(loc="upper right", fontsize=9)

    # 保存（処理後プロット）
    try:
        post_img = os.path.join(output_dir, f"{PREFIX}{base}_after.{IMG_FORMAT}")
        plt.tight_layout()
        plt.savefig(post_img, dpi=200)
        logging.info("Saved plot: %s", post_img)
    except Exception as e:
        logging.warning("Could not save plot %s: %s", post_img, e)

    plt.show()


if __name__ == "__main__":
    main()
