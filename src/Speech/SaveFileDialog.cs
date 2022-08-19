using Codeer.Friendly.Windows.Grasp;
using Codeer.Friendly.Windows.NativeStandardControls;

namespace Speech
{
    /// <summary>
    /// 名前を付けて保存ダイアログをラップします
    /// </summary>
    internal class SaveFileDialog
    {
        /// <summary>
        /// このクラスでラップしているWindowControlインスタンスを取得します
        /// </summary>
        public WindowControl Window { get; private set; }

        /// <summary>
        /// ファイル保存ダイアログを指定して初期化します
        /// </summary>
        /// <param name="dialog">名前を付けて保存ダイアログのインスタンス</param>
        public SaveFileDialog(WindowControl dialog)
        {
            Window = dialog;
        }

        /// <summary>
        /// ダイアログのファイル名を設定します
        /// </summary>
        /// <param name="path">保存先パス</param>
        public void SetFilePath(string path)
        {
            var combobox = Window.GetFromWindowClass("ComboBox");
            var textbox = new NativeEdit(combobox[combobox.Length - 1]); // 一番最後に取得できたものをファイルパス指定とする
            textbox.EmulateChangeText(path);
        }

        /// <summary>
        /// 保存ボタンを押します
        /// </summary>
        public void Save()
        {
            var save = new NativeButton(Window.IdentifyFromWindowText("保存(&S)")); // TODO: 日本語環境でしか動かない
            save.EmulateClick();
        }

        /// <summary>
        /// ファイル名を設定して保存ボタンを押します
        /// </summary>
        /// <param name="path">保存先パス</param>
        public void Save(string path)
        {
            SetFilePath(path);
            Save();
        }
    }
}
