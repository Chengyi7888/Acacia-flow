# Acacia Flow

Acacia Flow is a Windows desktop file conversion app built with Electron. It provides a clean drag-and-drop interface for converting common images, audio, video, PDFs, text files, and office-style documents locally on your computer.

## Download

Please download the latest Windows installer from the [GitHub Releases page](https://github.com/Chengyi7888/Acacia-flow/releases):

```text
Acacia.Flow.Setup.0.1.0.exe
```

After downloading, run the installer and launch Acacia Flow from the desktop shortcut or the Start Menu.

## Features

- Local file conversion with no cloud upload.
- Image conversion for common formats such as PNG, JPG, WEBP, GIF, AVIF, TIFF, BMP, and PDF.
- Audio conversion for MP3, WAV, FLAC, M4A, AAC, OGG, OPUS, and WMA.
- Video conversion for MP4, MOV, MKV, WEBM, GIF, and audio extraction.
- PDF conversion and text extraction for common PDF workflows.
- Text and table conversion for TXT, MD, HTML, CSV, RTF, DOCX, XLSX, and PDF.
- Conversion queue with removable items before conversion starts.
- Output folder selection, duplicate-name handling, and tray behavior settings.
- Animated assistant states that reflect file selection, conversion progress, completion, and errors.

## Offline Use

Acacia Flow runs locally after installation. The app does not require Codex, ChatGPT, local developer plugins, or a cloud service to operate.

FFmpeg and a portable LibreOffice runtime are bundled with the Windows installer. Advanced Office conversions such as DOC, PPT, and XLS workflows can run offline without asking the user to install additional plugins or applications.

The portable Office runtime is stored in `vendor/libreoffice` for source builds and is copied into the packaged application's resources during installation.

## Supported Conversion Examples

- Text to PDF, DOCX, XLSX, TXT, MD, HTML, CSV, and RTF.
- PDF to TXT, MD, HTML, DOCX, or cleaned PDF output.
- Image to PNG, JPG, WEBP, GIF, AVIF, TIFF, BMP, or PDF.
- Audio to MP3, WAV, FLAC, M4A, AAC, OGG, OPUS, or WMA.
- Video to MP4, MOV, MKV, WEBM, GIF, or common audio formats.
- Legacy Office documents such as DOC, PPT, and XLS to PDF or image formats when using the bundled offline runtime.

## Project Structure

```text
assets/                 App icons, background image, and assistant state images
scripts/                Build and smoke-test scripts
src/main/               Electron main process, preload script, and conversion engine
web/                    User interface
web/vendor/             Bundled browser-side libraries used by the app
vendor/libreoffice/     Bundled portable Office conversion runtime
package.json            App metadata and build configuration
```

## Privacy

Acacia Flow processes files on the local machine. User files are not uploaded by the app during normal conversion workflows.

## License

This project is released under the MIT License. Third-party dependencies remain under their own licenses.
