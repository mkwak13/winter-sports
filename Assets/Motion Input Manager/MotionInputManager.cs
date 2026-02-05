using System;
using System.Collections;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using UnityEngine;
using UnityEngine.Events;

/// <summary>
/// Class <c>MotionInputManager</c> handles the opening, changing and closing of the motion input window.
/// </summary>
public class MotionInputManager : MonoBehaviour
{
    // Window Dll constants
    private const UInt32 WM_CLOSE = 0x0010;
    private const int GWL_STYLE = -16;
    private const long WS_CAPTION = 0x00C00000L;
    private const long WS_BORDER = 0x00800000L;
    private const long WS_THICKFRAME = 0x00040000L;
    private readonly IntPtr HWND_TOPMOST = new(-1);

    // Windows DLL Functions
    [DllImport("user32.dll")] private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] private static extern long GetWindowLongA(IntPtr hWnd, int nIndex);
    [DllImport("user32.dll")] private static extern long SetWindowLongA(IntPtr hWnd, int nIndex, long dwNewLong);
    [DllImport("user32.dll")] private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int width, int height, uint uFlags);
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);

    // Motion Input Config
    private string controllerName = "Controller (XBOX 360 For Windows)";
    private string folder = "MotionInput";
    private string app = "motioninput.exe";
    private string configFile = "data/config.json";
    private string modesFolder = "data/modes";
    public string mode = "group6/superexplorers_racing";
    private string windowTitle = "Motioninput v3.4";

    [SerializeField] private RectTransform motionInputTransform;
    [SerializeField] private UnityEvent OnInitializedCallback; // Functions to be called when motion input is started
    [SerializeField] private bool activateOnStart = true; // Whether to start motion input when the scene is loaded
    [SerializeField] private bool keepPreviousMotionInput = false; // Whether to not terminate the previous motion input instance
    [SerializeField] private bool terminateWhenDestroyed = true; // Whether to destory the motion input instance when the scene is unloaded
    [SerializeField] private bool freezeGameWhenStarting = true; // Whether to freeze the game (stop everything that is running based on Time.timeScale)

    private IntPtr motionInputHWnd = IntPtr.Zero;
    private IntPtr unityWindowHandle;
    private bool isMotionInputStarted = false;
    private string configPath;

    // Start is called before the first frame update
    void Start()
    {
        // First we need to set up the config folder and move it over into persistentDataPath
        // persistentDataPath maps to AppData on Windows
        // can read and write to it
        // rest of MI will be stored in Application.streamingAssetsPath
        // because its read only

        // everything in Application.streamingAssetsPath at the start
        string readonly_configPath = Path.Combine(Application.streamingAssetsPath, folder, configFile);

        // copy over config.json to Application.persistentDataPath
        string configFolder = Path.Combine(Application.persistentDataPath, folder, "data");

        if (!Directory.Exists(configFolder))
        {
            Directory.CreateDirectory(configFolder);
        }

        // only read from that from now on
        configPath = Path.Combine(configFolder, "config.json");

        if (!File.Exists(configPath))
        {
            File.Copy(readonly_configPath, configPath);
        }


        Application.runInBackground = true;
        this.gameObject.SetActive(false);
        this.unityWindowHandle = GetForegroundWindow();
        if (this.activateOnStart)
        {
            this.StartMotionInput();
        }
    }

    // Called when the game object is destroyed
    void OnDestroy()
    {
        if (this.terminateWhenDestroyed)
        {
            this.TerminateMotionInput();
        }
    }

    // Called when the game is closed
    void OnApplicationQuit()
    {
        this.TerminateMotionInput();
    }

    /// <summary>
    /// Method <c>TerminateMotionInput</c> closes the motion input window.
    /// </summary>
    public void TerminateMotionInput()
    {
        IntPtr motionInputHWnd = FindWindow(null, this.windowTitle);
        if (motionInputHWnd == IntPtr.Zero) return;
        SendMessage(motionInputHWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
    }

    /// <summary>
    /// Method <c>OpenMotionInput</c> opens the motion input window using powershell
    /// </summary>
    private void OpenMotionInput()
    {
        TerminateMotionInput();
        // Start Motion Input Window using powershell
        string appPath = Path.Combine(Application.streamingAssetsPath, folder);
        string processArgs = $"--config \\\"{configPath}\\\"";

        // 2. Wrap the executable name in single quotes.
        string quotedApp = $"'{app}'";
        string psCommand = $"& {{ cd '{appPath}'; Start-Process {quotedApp} -ArgumentList @('{processArgs}') }}";
        ProcessStartInfo startInfo = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = $"-NoProfile -WindowStyle Hidden -Command \"{psCommand}\"",

            // string : cd "<appPath>"; Start-Process "<app>" -ArgumentList "--config `"<path>`""
            UseShellExecute = false,
            CreateNoWindow = true,
        };
        try
        {
            using (Process process = new Process())
            {
                process.StartInfo = startInfo;
                process.Start();
            }
        }
        catch (Exception ex)
        {
            UnityEngine.Debug.LogError($"An error occurred: {ex.Message}");
        }
    }

    /// <summary>
    /// Method <c>RemoveMotionInputHeader</c> removes the title bar (including the close, maximize and minimize button) of the motion input window.
    /// </summary>
    private void RemoveMotionInputHeader()
    {
        if (this.motionInputHWnd == IntPtr.Zero) return;
        long windowStyle = GetWindowLongA(this.motionInputHWnd, GWL_STYLE);
        windowStyle &= ~(WS_BORDER | WS_CAPTION | WS_THICKFRAME);
        SetWindowLongA(this.motionInputHWnd, GWL_STYLE, windowStyle);
    }

    /// <summary>
    /// Method <c>SetMotionInputWindowTransform</c> sets the position and size of the motion input window based on the window preview in prefab.
    /// </summary>
    private void SetMotionInputWindowTransform()
    {
        // Check if motion input instance exists
        if (this.motionInputHWnd == IntPtr.Zero) return;
        // Get the window preview game object corners
        Vector3[] corners = new Vector3[4];
        this.motionInputTransform.GetWorldCorners(corners);
        Vector3 screenPos = corners[1];
        screenPos.y = Screen.height - screenPos.y;
        Vector2Int position = Vector2Int.RoundToInt(screenPos);
        float width = Vector3.Distance(corners[0], corners[3]);
        float height = Vector3.Distance(corners[0], corners[1]);
        Vector2Int dimension = Vector2Int.RoundToInt(new Vector2(width, height));
        // Set the motion input window position and size
        SetWindowPos(this.motionInputHWnd, HWND_TOPMOST, position.x, position.y, dimension.x, dimension.y, 0x0000);
    }

    /// <summary>
    /// Method <c>DetectMotionInputWindowStarted</c> detects whether there is an instance of the motion input window.
    /// </summary>
    /// <returns></returns>
    private IEnumerator DetectMotionInputWindowStarted()
    {
        UnityEngine.Debug.Log("Starting MotionInput window");
        // Wait until the motion input instance exists
        while (this.motionInputHWnd == IntPtr.Zero)
        {
            this.motionInputHWnd = FindWindow(null, windowTitle);
            yield return new WaitForSecondsRealtime(0.5f); // Detect every 0.5 seconds
        }
        UnityEngine.Debug.Log("MotionInput window started");
        this.PostStartUpHandler();
        StartCoroutine(DetectMotionInputInitialized());
    }

    /// <summary>
    /// Method <c>DetectMotionInputInitialized</c> detects whether a controller from motion input window is connected to Unity.
    /// </summary>
    /// <returns></returns>
    private IEnumerator DetectMotionInputInitialized()
    {
        yield return null;
        UnityEngine.Debug.Log("MotionInput initialized");
        if (this.freezeGameWhenStarting)
        {
            Time.timeScale = 1;
        }

        // Notify SkiScoreManager to show UI
        if (SkiScoreManager.Instance != null)
        {
            SkiScoreManager.Instance.SetUIVisibility(true);
        }

        this.gameObject.SetActive(false);
        if (this.OnInitializedCallback != null)
        {
            this.OnInitializedCallback.Invoke();
        }
    }

    /// <summary>
    /// Method <c>PostStartUpHandler</c> handles setting the styles and position of the motion input window.
    /// </summary>
    private void PostStartUpHandler()
    {
        this.RemoveMotionInputHeader();
        this.SetMotionInputWindowTransform();
        SetForegroundWindow(this.unityWindowHandle);
    }

    /// <summary>
    /// Method <c>Change Mode</c> edits the json file with the new mode.
    /// </summary>
    /// <returns></returns>
    private bool ChangeMode()
    {
        string modePath = Path.Combine(Application.streamingAssetsPath, folder, modesFolder, this.mode + ".json");
        if (!File.Exists(modePath))
        {
            UnityEngine.Debug.LogError($"Mode {this.mode} not found");
            return false;
        }

        // Write to config.json in persistantDataPath
        string[] jsonString = File.ReadAllLines(configPath);

        string newModeString = $"\"mode\": \"{this.mode}\",";

        jsonString[1] = newModeString;

        File.WriteAllLines(configPath, jsonString);
        return true;
    }

    /// <summary>
    /// Method <c>GetCurrentMode</c> gets the mode inside the config.json file
    /// </summary>
    /// <returns></returns>
    private string GetCurrentMode()
    {
        // Read from config.json
        string[] jsonString = File.ReadAllLines(configPath);
        string modeValue = jsonString[1].Split(":")[1];

        return modeValue.Trim().Substring(1, modeValue.Length - 4);
    }

    /// <summary>
    /// Method <c>StartMotionInput</c> handles the opening of the motion input window.
    /// </summary>
    private void StartMotionInput()
    {
        this.gameObject.SetActive(true);
        this.mode = this.mode.Replace(".json", ""); // Remove duplicate .json file extension
        this.motionInputHWnd = FindWindow(null, this.windowTitle); // Check if previous motion input instance exists
        // Check if the conditions for keep previous instance
        if (this.keepPreviousMotionInput && this.GetCurrentMode().Equals(this.mode) && this.motionInputHWnd != IntPtr.Zero)
        {
            UnityEngine.Debug.Log("Previous instance found");
            this.isMotionInputStarted = true;
            if (this.OnInitializedCallback != null)
            {
                this.OnInitializedCallback.Invoke();
            }
            UnityEngine.Debug.Log("MotionInput initialized");
            this.gameObject.SetActive(false);
            return;
        }
        if (this.freezeGameWhenStarting) Time.timeScale = 0; // Freeze game
        if (!this.ChangeMode()) return; // Check if mode is valid, else return
        Thread motionInputThread = new Thread(this.OpenMotionInput);
        motionInputThread.Start();
        StartCoroutine(DetectMotionInputWindowStarted());
    }

    /// <summary>
    /// Method <c>IsMotionInputInitialized</c> is a getter to check whether the motion input window has started.
    /// </summary>
    /// <returns>true for started, false for not started</returns>
    public bool IsMotionInputInitialized()
    {
        return this.isMotionInputStarted;
    }
}