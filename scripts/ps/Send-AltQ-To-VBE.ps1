[CmdletBinding()]
param(
	[ValidateSet("Q", "F4")]
	[string]$Key = "Q"
)

$ErrorActionPreference = "Stop"

Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class VbeWindowApi
{
	public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

	[DllImport("user32.dll")]
	public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

	[DllImport("user32.dll")]
	public static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int maxCount);

	[DllImport("user32.dll")]
	public static extern bool IsWindowVisible(IntPtr hwnd);

	[DllImport("user32.dll")]
	public static extern bool ShowWindow(IntPtr hwnd, int command);

	[DllImport("user32.dll")]
	public static extern bool BringWindowToTop(IntPtr hwnd);

	[DllImport("user32.dll")]
	public static extern bool SetForegroundWindow(IntPtr hwnd);

	[DllImport("user32.dll")]
	public static extern IntPtr GetForegroundWindow();

	[DllImport("user32.dll")]
	public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);

	[DllImport("kernel32.dll")]
	public static extern uint GetCurrentThreadId();

	[DllImport("user32.dll")]
	public static extern bool AttachThreadInput(uint sourceThreadId, uint targetThreadId, bool attach);

	[StructLayout(LayoutKind.Sequential)]
	private struct INPUT
	{
		public uint type;
		public INPUTUNION input;
	}

	[StructLayout(LayoutKind.Explicit)]
	private struct INPUTUNION
	{
		[FieldOffset(0)] public MOUSEINPUT mouse;
		[FieldOffset(0)] public KEYBDINPUT keyboard;
		[FieldOffset(0)] public HARDWAREINPUT hardware;
	}

	[StructLayout(LayoutKind.Sequential)]
	private struct MOUSEINPUT
	{
		public int x;
		public int y;
		public uint mouseData;
		public uint flags;
		public uint time;
		public UIntPtr extraInfo;
	}

	[StructLayout(LayoutKind.Sequential)]
	private struct KEYBDINPUT
	{
		public ushort virtualKey;
		public ushort scanCode;
		public uint flags;
		public uint time;
		public UIntPtr extraInfo;
	}

	[StructLayout(LayoutKind.Sequential)]
	private struct HARDWAREINPUT
	{
		public uint message;
		public ushort parameterLow;
		public ushort parameterHigh;
	}

	[DllImport("user32.dll", SetLastError = true)]
	private static extern uint SendInput(uint count, INPUT[] inputs, int size);

	public static IntPtr FindVbeWindow()
	{
		IntPtr result = IntPtr.Zero;
		EnumWindows((hwnd, lParam) =>
		{
			var title = new StringBuilder(256);
			GetWindowText(hwnd, title, title.Capacity);
			if (IsWindowVisible(hwnd) && title.ToString().Contains("Microsoft Visual Basic"))
			{
				result = hwnd;
				return false;
			}
			return true;
		}, IntPtr.Zero);
		return result;
	}

	public static bool SendAltQ()
	{
		const uint inputKeyboard = 1;
		const uint keyUp = 2;
		const ushort virtualKeyMenu = 0x12;
		const ushort virtualKeyQ = 0x51;
		var inputs = new[]
		{
			new INPUT { type = inputKeyboard, input = new INPUTUNION { keyboard = new KEYBDINPUT { virtualKey = virtualKeyMenu } } },
			new INPUT { type = inputKeyboard, input = new INPUTUNION { keyboard = new KEYBDINPUT { virtualKey = virtualKeyQ } } },
			new INPUT { type = inputKeyboard, input = new INPUTUNION { keyboard = new KEYBDINPUT { virtualKey = virtualKeyQ, flags = keyUp } } },
			new INPUT { type = inputKeyboard, input = new INPUTUNION { keyboard = new KEYBDINPUT { virtualKey = virtualKeyMenu, flags = keyUp } } }
		};
		return SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT))) == inputs.Length;
	}

	public static bool SendAltF4()
	{
		const uint inputKeyboard = 1;
		const uint keyUp = 2;
		const ushort virtualKeyMenu = 0x12;
		const ushort virtualKeyF4 = 0x73;
		var inputs = new[]
		{
			new INPUT { type = inputKeyboard, input = new INPUTUNION { keyboard = new KEYBDINPUT { virtualKey = virtualKeyMenu } } },
			new INPUT { type = inputKeyboard, input = new INPUTUNION { keyboard = new KEYBDINPUT { virtualKey = virtualKeyF4 } } },
			new INPUT { type = inputKeyboard, input = new INPUTUNION { keyboard = new KEYBDINPUT { virtualKey = virtualKeyF4, flags = keyUp } } },
			new INPUT { type = inputKeyboard, input = new INPUTUNION { keyboard = new KEYBDINPUT { virtualKey = virtualKeyMenu, flags = keyUp } } }
		};
		return SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT))) == inputs.Length;
	}
}
"@

$vbeHwnd = [VbeWindowApi]::FindVbeWindow()
if ($vbeHwnd -eq [IntPtr]::Zero) {
	throw "No visible VBE window was found."
}

$processId = [uint32]0
$vbeThreadId = [VbeWindowApi]::GetWindowThreadProcessId($vbeHwnd, [ref]$processId)
$currentThreadId = [VbeWindowApi]::GetCurrentThreadId()
$foregroundBefore = [VbeWindowApi]::GetForegroundWindow()
$foregroundProcessId = [uint32]0
$foregroundThreadId = [VbeWindowApi]::GetWindowThreadProcessId($foregroundBefore, [ref]$foregroundProcessId)
$attached = $false

try {
	if ($foregroundThreadId -and $foregroundThreadId -ne $currentThreadId) {
		$attached = [VbeWindowApi]::AttachThreadInput($currentThreadId, $foregroundThreadId, $true)
	}

	[void][VbeWindowApi]::ShowWindow($vbeHwnd, 9) # SW_RESTORE
	[void][VbeWindowApi]::BringWindowToTop($vbeHwnd)
	if (-not [VbeWindowApi]::SetForegroundWindow($vbeHwnd)) {
		throw "SetForegroundWindow failed for VBE HWND 0x$($vbeHwnd.ToString('X8'))."
	}
} finally {
	if ($attached) {
		[void][VbeWindowApi]::AttachThreadInput($currentThreadId, $foregroundThreadId, $false)
	}
}

Start-Sleep -Milliseconds 300
$foreground = [VbeWindowApi]::GetForegroundWindow()
if ($foreground -ne $vbeHwnd) {
	throw "The VBE was not foreground. Expected 0x$($vbeHwnd.ToString('X8')), got 0x$($foreground.ToString('X8'))."
}

$sendSucceeded = if ($Key -eq "F4") {
	Write-Host ("Sending Alt+F4 to VBE HWND 0x{0:X8} (PID {1})..." -f $vbeHwnd.ToInt64(), $processId) -ForegroundColor Cyan
	[VbeWindowApi]::SendAltF4()
} else {
	Write-Host ("Sending Alt+Q to VBE HWND 0x{0:X8} (PID {1})..." -f $vbeHwnd.ToInt64(), $processId) -ForegroundColor Cyan
	[VbeWindowApi]::SendAltQ()
}
if (-not $sendSucceeded) {
	throw "SendInput failed while sending Alt+$Key."
}

Start-Sleep -Milliseconds 500
Write-Host ("VBE window still visible: {0}" -f [VbeWindowApi]::IsWindowVisible($vbeHwnd)) -ForegroundColor Yellow
