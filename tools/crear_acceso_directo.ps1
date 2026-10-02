# Crea "Compilador de reportes APEX.lnk" en la raiz del proyecto.
#
# Un acceso directo normal guarda una ruta absoluta y deja de funcionar si la
# carpeta se mueve o se entrega a otra persona. Este acceso directo apunta a
# powershell.exe (presente en todo Windows) con un lanzador minimo: Windows
# entrega a la aplicacion la ruta del propio .lnk en STARTUPINFO.lpTitle, el
# lanzador toma esa carpeta y ejecuta "Abrir compilador.bat" que esta a su lado.
# Asi funciona en cualquier ubicacion, incluida una copia comprimida.
#
# Se eliminan ademas los bloques de seguimiento de vinculos y de propiedades,
# que contienen el nombre del equipo y el SID del usuario que lo creo.
#
# El acceso directo debe quedar en la misma carpeta que el .bat; no funciona
# copiado al Escritorio ni abierto dentro de un ZIP sin extraer.
#
# Uso (desde la raiz del proyecto):
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\crear_acceso_directo.ps1
param(
    [string]$Folder = (Split-Path -Parent $PSScriptRoot),
    [string]$Target = 'Abrir compilador.bat',
    [string]$Name = 'Compilador de reportes APEX.lnk'
)

$ErrorActionPreference = 'Stop'
$Folder = (Resolve-Path -LiteralPath $Folder).Path
$path = Join-Path $Folder $Name

$launcher = @"
Add-Type -Name K -Namespace W -Mem '[DllImport("kernel32")]public static extern void GetStartupInfoW(IntPtr p);'
`$m=[Runtime.InteropServices.Marshal];`$p=`$m::AllocHGlobal(99);[W.K]::GetStartupInfoW(`$p)
`$d=Split-Path(`$m::PtrToStringUni(`$m::ReadIntPtr(`$p,3*[IntPtr]::Size)))
saps "`$d\$Target" -wo `$d -wi Mi
"@
$arguments = '-NoP -NonI -W Hidden -Enc ' + [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($launcher))
if ($arguments.Length -gt 1023) { throw 'El lanzador supera el limite de 1023 caracteres de WScript.' }

$shell = New-Object -ComObject WScript.Shell
$link = $shell.CreateShortcut($path)
$link.TargetPath = '%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe'
$link.Arguments = $arguments
$link.WorkingDirectory = ''
$link.WindowStyle = 7
$link.IconLocation = '%SystemRoot%\System32\imageres.dll,109'
$link.Description = 'Abre el compilador local de plantillas Word para Oracle APEX'
$link.Save()

# Quita TrackerDataBlock (0xA0000003) y PropertyStoreDataBlock (0xA0000009),
# MS-SHLLINK 2.5.10 y 2.5.7.
$bytes = [IO.File]::ReadAllBytes($path)
$flags = [BitConverter]::ToUInt32($bytes, 0x14)
$offset = 0x4C
if ($flags -band 0x1) { $offset += 2 + [BitConverter]::ToUInt16($bytes, $offset) }
if ($flags -band 0x2) { $offset += [BitConverter]::ToUInt32($bytes, $offset) }
foreach ($bit in 0x4, 0x8, 0x10, 0x20, 0x40) {
    if ($flags -band $bit) {
        $count = [BitConverter]::ToUInt16($bytes, $offset)
        $offset += 2 + $(if ($flags -band 0x80) { 2 * $count } else { $count })
    }
}
$output = New-Object System.Collections.Generic.List[byte]
$output.AddRange([byte[]]$bytes[0..($offset - 1)])
while ($offset + 4 -le $bytes.Length) {
    $size = [BitConverter]::ToUInt32($bytes, $offset)
    if ($size -lt 4) { break }
    $signature = [BitConverter]::ToUInt32($bytes, $offset + 4)
    # Firmas como UInt32: el literal hexadecimal de PowerShell seria Int32 negativo.
    if ($signature -ne [uint32]2684354563 -and $signature -ne [uint32]2684354569) {
        $output.AddRange([byte[]]$bytes[$offset..($offset + $size - 1)])
    }
    $offset += $size
}
$output.AddRange([byte[]](0, 0, 0, 0))
[IO.File]::WriteAllBytes($path, $output.ToArray())
Write-Output $path
