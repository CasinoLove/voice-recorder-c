/*
CasinoLove Voice Recorder for Windows
Lightweight native voice recorder written in C.

Created by CasinoLove (Casinolove Kft.)

Official company details:
Registered name: Casinolove Kft.
Jurisdiction: Hungary (European Union)
Company Registration Number (Hungary): 14-09-318400
Email: hello@casinolove.org
Website: https://hu.casinolove.org/
GitHub: https://github.com/CasinoLove

Repository:
https://github.com/CasinoLove/voice-recorder-c

*/


#define _CRT_SECURE_NO_WARNINGS
#include <windows.h>
#include <mmsystem.h>
#include <conio.h>
#include <io.h>
#include <stdio.h>
#include <stdlib.h>
#include <stdint.h>
#include <string.h>

#define INPUT_BUFFER_COUNT 4
#define INPUT_BUFFER_BYTES 8192
#define MAX_OUTPUT_QUEUE 32
#define HISTORY_COLS 60
#define HISTORY_ROWS 10
#define FILE_SYNC_INTERVAL_MS 1000

typedef struct AppSettings {
    char outputPath[MAX_PATH];
    int inputUseMapper;
    UINT inputDeviceId;
    int outputUseMapper;
    UINT outputDeviceId;
    int monitorEnabled;
    DWORD sampleRate;
    WORD channels;
    DWORD refreshIntervalMs;
} AppSettings;

typedef struct HistoryEntry {
    ULONGLONG slotIndex;
    LONG peak;
} HistoryEntry;

typedef struct WavHeader {
    char riff[4];
    uint32_t overall_size;
    char wave[4];
    char fmt_chunk_marker[4];
    uint32_t length_of_fmt;
    uint16_t format_type;
    uint16_t channels;
    uint32_t sample_rate;
    uint32_t byterate;
    uint16_t block_align;
    uint16_t bits_per_sample;
    char data_chunk_header[4];
    uint32_t data_size;
} WavHeader;

typedef struct OutputBlock {
    WAVEHDR hdr;
    BYTE* data;
} OutputBlock;

typedef struct AppState {
    AppSettings settings;
    HWAVEIN hWaveIn;
    HWAVEOUT hWaveOut;
    FILE* outFile;
    volatile LONG running;
    volatile LONG lastPeak;
    volatile LONG currentBinPeak;
    uint32_t dataBytesWritten;
    WAVEFORMATEX fmt;
    CRITICAL_SECTION lock;
    char inputDeviceName[256];
    char outputDeviceName[256];
    ULONGLONG startTickMs;
} AppState;

static const DWORD kRates[3] = { 44100, 48000, 192000 };
static const DWORD kRefreshChoices[5] = { 25, 50, 100, 200, 500 };

static AppState g_app = {0};
static OutputBlock* g_outputBlocks[MAX_OUTPUT_QUEUE] = {0};
static volatile LONG g_outputCount = 0;

static void format_uint_with_commas(uint32_t value, char* out, size_t outSize) {
    char tmp[32];
    int len;
    int commas;
    int firstGroup;
    int i, j = 0;

    snprintf(tmp, sizeof(tmp), "%u", value);
    len = (int)strlen(tmp);
    commas = (len - 1) / 3;
    firstGroup = len % 3;
    if (firstGroup == 0) firstGroup = 3;

    for (i = 0; i < len && j + 1 < (int)outSize; ++i) {
        out[j++] = tmp[i];
        if (i == len - 1) {
            continue;
        }
        if (i < firstGroup - 1) {
            continue;
        }
        if (((i - (firstGroup - 1)) % 3) == 0) {
            if (commas > 0 && j + 1 < (int)outSize) {
                out[j++] = ',';
                commas--;
            }
        }
    }
    out[j] = '\0';
}


static void clear_console_screen(void) {
    HANDLE hConsole;
    CONSOLE_SCREEN_BUFFER_INFO csbi;
    DWORD count;
    DWORD cellCount;
    COORD homeCoords;

    hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
    homeCoords.X = 0;
    homeCoords.Y = 0;

    if (hConsole == INVALID_HANDLE_VALUE) return;
    if (!GetConsoleScreenBufferInfo(hConsole, &csbi)) return;

    cellCount = (DWORD)(csbi.dwSize.X * csbi.dwSize.Y);
    FillConsoleOutputCharacterA(hConsole, ' ', cellCount, homeCoords, &count);
    FillConsoleOutputAttribute(hConsole, csbi.wAttributes, cellCount, homeCoords, &count);
    SetConsoleCursorPosition(hConsole, homeCoords);
}

static void move_cursor_home(void) {
    COORD c;
    c.X = 0;
    c.Y = 0;
    SetConsoleCursorPosition(GetStdHandle(STD_OUTPUT_HANDLE), c);
}

static void strip_newline(char* s) {
    size_t len;
    if (!s) return;
    len = strlen(s);
    while (len > 0 && (s[len - 1] == '\n' || s[len - 1] == '\r')) {
        s[len - 1] = '\0';
        len--;
    }
}

static void read_line(const char* prompt, char* out, size_t outSize) {
    if (!out || outSize == 0) return;
    printf("%s", prompt);
    fflush(stdout);
    if (!fgets(out, (int)outSize, stdin)) {
        out[0] = '\0';
        return;
    }
    strip_newline(out);
}

static int ask_yes_no(const char* prompt, int defaultYes) {
    char line[32];
    read_line(prompt, line, sizeof(line));
    if (line[0] == '\0') return defaultYes;
    if (line[0] == 'y' || line[0] == 'Y') return 1;
    if (line[0] == 'n' || line[0] == 'N') return 0;
    return defaultYes;
}

static int read_menu_int(const char* prompt, int minValue, int maxValue, int* outValue) {
    char line[64];
    char* endPtr;
    long v;

    read_line(prompt, line, sizeof(line));
    if (line[0] == '\0') return 0;

    v = strtol(line, &endPtr, 10);
    if (*endPtr != '\0') return 0;
    if (v < minValue || v > maxValue) return 0;
    *outValue = (int)v;
    return 1;
}

static void wait_for_enter(void) {
    char line[8];
    printf("\nPress ENTER to continue...");
    fflush(stdout);
    fgets(line, sizeof(line), stdin);
}

static int file_exists(const char* path) {
    DWORD attr;
    if (!path || !path[0]) return 0;
    attr = GetFileAttributesA(path);
    return attr != INVALID_FILE_ATTRIBUTES && !(attr & FILE_ATTRIBUTE_DIRECTORY);
}

static const char* channels_name(WORD channels) {
    return channels == 1 ? "Mono" : "Stereo";
}

static void get_input_device_name(int useMapper, UINT deviceId, char* out, size_t outSize) {
    WAVEINCAPSA caps;
    MMRESULT mmr;

    if (useMapper) {
        snprintf(out, outSize, "Default system input");
        return;
    }

    mmr = waveInGetDevCapsA(deviceId, &caps, sizeof(caps));
    if (mmr == MMSYSERR_NOERROR) {
        snprintf(out, outSize, "%s", caps.szPname);
    } else {
        snprintf(out, outSize, "Input device %u", deviceId);
    }
}

static void get_output_device_name(int useMapper, UINT deviceId, char* out, size_t outSize) {
    WAVEOUTCAPSA caps;
    MMRESULT mmr;

    if (useMapper) {
        snprintf(out, outSize, "Default system output");
        return;
    }

    mmr = waveOutGetDevCapsA(deviceId, &caps, sizeof(caps));
    if (mmr == MMSYSERR_NOERROR) {
        snprintf(out, outSize, "%s", caps.szPname);
    } else {
        snprintf(out, outSize, "Output device %u", deviceId);
    }
}

static void build_format(DWORD sampleRate, WORD channels, WAVEFORMATEX* fmt) {
    memset(fmt, 0, sizeof(*fmt));
    fmt->wFormatTag = WAVE_FORMAT_PCM;
    fmt->nChannels = channels;
    fmt->nSamplesPerSec = sampleRate;
    fmt->wBitsPerSample = 16;
    fmt->nBlockAlign = (WORD)((channels * 16) / 8);
    fmt->nAvgBytesPerSec = fmt->nSamplesPerSec * fmt->nBlockAlign;
    fmt->cbSize = 0;
}

static int query_input_support_internal(int useMapper, UINT deviceId, const WAVEFORMATEX* fmt) {
    HWAVEIN hwi = NULL;
    MMRESULT mmr;
    UINT openId = useMapper ? WAVE_MAPPER : deviceId;

    mmr = waveInOpen(&hwi, openId, (WAVEFORMATEX*)fmt, 0, 0, WAVE_FORMAT_QUERY);
    if (mmr == MMSYSERR_NOERROR && hwi) waveInClose(hwi);
    return mmr == MMSYSERR_NOERROR;
}

static int query_output_support_internal(int useMapper, UINT deviceId, const WAVEFORMATEX* fmt) {
    HWAVEOUT hwo = NULL;
    MMRESULT mmr;
    UINT openId = useMapper ? WAVE_MAPPER : deviceId;

    mmr = waveOutOpen(&hwo, openId, (WAVEFORMATEX*)fmt, 0, 0, WAVE_FORMAT_QUERY);
    if (mmr == MMSYSERR_NOERROR && hwo) waveOutClose(hwo);
    return mmr == MMSYSERR_NOERROR;
}

static int combo_supported(const AppSettings* settings, DWORD rate, WORD channels) {
    WAVEFORMATEX fmt;
    int ok;

    build_format(rate, channels, &fmt);
    ok = query_input_support_internal(settings->inputUseMapper, settings->inputDeviceId, &fmt);
    if (!ok) return 0;

    if (settings->monitorEnabled) {
        ok = query_output_support_internal(settings->outputUseMapper, settings->outputDeviceId, &fmt);
        if (!ok) return 0;
    }

    return 1;
}

static int current_combo_supported(const AppSettings* settings) {
    return combo_supported(settings, settings->sampleRate, settings->channels);
}

static void choose_best_defaults(AppSettings* settings) {
    if (combo_supported(settings, 48000, 1)) {
        settings->sampleRate = 48000;
        settings->channels = 1;
        return;
    }
    if (combo_supported(settings, 44100, 1)) {
        settings->sampleRate = 44100;
        settings->channels = 1;
        return;
    }
    if (combo_supported(settings, 192000, 1)) {
        settings->sampleRate = 192000;
        settings->channels = 1;
        return;
    }
    if (combo_supported(settings, 48000, 2)) {
        settings->sampleRate = 48000;
        settings->channels = 2;
        return;
    }
    if (combo_supported(settings, 44100, 2)) {
        settings->sampleRate = 44100;
        settings->channels = 2;
        return;
    }
    if (combo_supported(settings, 192000, 2)) {
        settings->sampleRate = 192000;
        settings->channels = 2;
        return;
    }

    settings->sampleRate = 48000;
    settings->channels = 1;
}

static void adjust_to_supported_combo(AppSettings* settings) {
    if (!current_combo_supported(settings)) choose_best_defaults(settings);
}

static int write_placeholder_header(FILE* f, const WAVEFORMATEX* fmt) {
    WavHeader h;
    memcpy(h.riff, "RIFF", 4);
    h.overall_size = 0;
    memcpy(h.wave, "WAVE", 4);
    memcpy(h.fmt_chunk_marker, "fmt ", 4);
    h.length_of_fmt = 16;
    h.format_type = 1;
    h.channels = fmt->nChannels;
    h.sample_rate = fmt->nSamplesPerSec;
    h.byterate = fmt->nAvgBytesPerSec;
    h.block_align = fmt->nBlockAlign;
    h.bits_per_sample = fmt->wBitsPerSample;
    memcpy(h.data_chunk_header, "data", 4);
    h.data_size = 0;
    return fwrite(&h, sizeof(h), 1, f) == 1 ? 1 : 0;
}

static int update_wav_header(FILE* f, uint32_t dataBytesWritten, const WAVEFORMATEX* fmt) {
    WavHeader h;
    long pos;

    pos = ftell(f);
    if (pos < 0) pos = 0;

    memcpy(h.riff, "RIFF", 4);
    h.overall_size = 36 + dataBytesWritten;
    memcpy(h.wave, "WAVE", 4);
    memcpy(h.fmt_chunk_marker, "fmt ", 4);
    h.length_of_fmt = 16;
    h.format_type = 1;
    h.channels = fmt->nChannels;
    h.sample_rate = fmt->nSamplesPerSec;
    h.byterate = fmt->nAvgBytesPerSec;
    h.block_align = fmt->nBlockAlign;
    h.bits_per_sample = fmt->wBitsPerSample;
    memcpy(h.data_chunk_header, "data", 4);
    h.data_size = dataBytesWritten;

    if (fseek(f, 0, SEEK_SET) != 0) return 0;
    if (fwrite(&h, sizeof(h), 1, f) != 1) return 0;
    fflush(f);
    if (fseek(f, pos, SEEK_SET) != 0) return 0;
    return 1;
}

static LONG compute_peak_16bit_pcm(const BYTE* data, DWORD bytesRecorded, int channels) {
    const int16_t* samples = (const int16_t*)data;
    DWORD sampleCount = bytesRecorded / sizeof(int16_t);
    LONG peak = 0;
    DWORD i;
    (void)channels;

    for (i = 0; i < sampleCount; ++i) {
        LONG v = samples[i];
        if (v == -32768) v = 32767;
        else if (v < 0) v = -v;
        if (v > peak) peak = v;
    }
    return peak;
}

static void cleanup_finished_output_blocks(void) {
    int i;
    for (i = 0; i < MAX_OUTPUT_QUEUE; ++i) {
        OutputBlock* block = g_outputBlocks[i];
        if (block && (block->hdr.dwFlags & WHDR_DONE)) {
            waveOutUnprepareHeader(g_app.hWaveOut, &block->hdr, sizeof(block->hdr));
            free(block->data);
            free(block);
            g_outputBlocks[i] = NULL;
            InterlockedDecrement(&g_outputCount);
        }
    }
}

static void enqueue_monitor_audio(const BYTE* data, DWORD bytesRecorded) {
    OutputBlock* block;
    int i;

    if (!g_app.settings.monitorEnabled || !g_app.hWaveOut || bytesRecorded == 0) return;

    cleanup_finished_output_blocks();
    if (g_outputCount >= MAX_OUTPUT_QUEUE - 1) return;

    block = (OutputBlock*)calloc(1, sizeof(OutputBlock));
    if (!block) return;

    block->data = (BYTE*)malloc(bytesRecorded);
    if (!block->data) {
        free(block);
        return;
    }

    memcpy(block->data, data, bytesRecorded);
    block->hdr.lpData = (LPSTR)block->data;
    block->hdr.dwBufferLength = bytesRecorded;

    if (waveOutPrepareHeader(g_app.hWaveOut, &block->hdr, sizeof(block->hdr)) != MMSYSERR_NOERROR) {
        free(block->data);
        free(block);
        return;
    }

    if (waveOutWrite(g_app.hWaveOut, &block->hdr, sizeof(block->hdr)) != MMSYSERR_NOERROR) {
        waveOutUnprepareHeader(g_app.hWaveOut, &block->hdr, sizeof(block->hdr));
        free(block->data);
        free(block);
        return;
    }

    for (i = 0; i < MAX_OUTPUT_QUEUE; ++i) {
        if (!g_outputBlocks[i]) {
            g_outputBlocks[i] = block;
            InterlockedIncrement(&g_outputCount);
            return;
        }
    }

    while (!(block->hdr.dwFlags & WHDR_DONE)) Sleep(1);
    waveOutUnprepareHeader(g_app.hWaveOut, &block->hdr, sizeof(block->hdr));
    free(block->data);
    free(block);
}

static void sync_recording_file_locked(void) {
    if (!g_app.outFile) return;
    fflush(g_app.outFile);
    update_wav_header(g_app.outFile, g_app.dataBytesWritten, &g_app.fmt);
    fflush(g_app.outFile);
    _commit(_fileno(g_app.outFile));
}

static void CALLBACK wave_in_proc(HWAVEIN hwi, UINT msg, DWORD_PTR instance, DWORD_PTR param1, DWORD_PTR param2) {
    WAVEHDR* hdr;
    LONG peak;

    (void)hwi;
    (void)instance;
    (void)param2;

    if (msg != WIM_DATA) return;

    hdr = (WAVEHDR*)param1;
    if (!hdr || hdr->dwBytesRecorded == 0) return;

    if (InterlockedCompareExchange(&g_app.running, 0, 0)) {
        EnterCriticalSection(&g_app.lock);

        if (g_app.outFile) {
            fwrite(hdr->lpData, 1, hdr->dwBytesRecorded, g_app.outFile);
            g_app.dataBytesWritten += hdr->dwBytesRecorded;
        }

        peak = compute_peak_16bit_pcm((const BYTE*)hdr->lpData, hdr->dwBytesRecorded, g_app.fmt.nChannels);
        g_app.lastPeak = peak;
        if (peak > g_app.currentBinPeak) {
            g_app.currentBinPeak = peak;
        }

        if (g_app.settings.monitorEnabled) {
            enqueue_monitor_audio((const BYTE*)hdr->lpData, hdr->dwBytesRecorded);
        }

        LeaveCriticalSection(&g_app.lock);

        hdr->dwBytesRecorded = 0;
        waveInAddBuffer(g_app.hWaveIn, hdr, sizeof(*hdr));
    }
}

static BOOL WINAPI console_ctrl_handler(DWORD ctrlType) {
    switch (ctrlType) {
        case CTRL_C_EVENT:
        case CTRL_BREAK_EVENT:
        case CTRL_CLOSE_EVENT:
            InterlockedExchange(&g_app.running, 0);
            return TRUE;
        default:
            return FALSE;
    }
}

static void seconds_to_hms(ULONGLONG totalSeconds, int* h, int* m, int* s) {
    *h = (int)(totalSeconds / 3600ULL);
    *m = (int)((totalSeconds % 3600ULL) / 60ULL);
    *s = (int)(totalSeconds % 60ULL);
}

static char chart_char_for_cell(const HistoryEntry* entry, int rowIndex) {
    double pct = (entry->peak / 32767.0) * 100.0;
    double threshold = (double)(HISTORY_ROWS - rowIndex) * 10.0;

    if (pct >= threshold) {
        if (pct > 80.0 && rowIndex <= 1) return '!';
        return '#';
    }

    if (rowIndex == HISTORY_ROWS - 1 && pct < 20.0) return '_';
    return ' ';
}

static void render_history_chart(const HistoryEntry* history, int historyCount) {
    int row, col;
    int start = historyCount > HISTORY_COLS ? historyCount - HISTORY_COLS : 0;
    int shown = historyCount - start;

    for (row = 0; row < HISTORY_ROWS; ++row) {
        int label = (HISTORY_ROWS - row) * 10;
        printf("%3d|", label);
        for (col = 0; col < HISTORY_COLS; ++col) {
            if (col < HISTORY_COLS - shown) {
                putchar(' ');
            } else {
                int idx = start + (col - (HISTORY_COLS - shown));
                putchar(chart_char_for_cell(&history[idx], row));
            }
        }
        printf("|\n");
    }

    printf("   +");
    for (col = 0; col < HISTORY_COLS; ++col) putchar('-');
    printf("+\n");
}

static void render_session_screen(const HistoryEntry* history, int historyCount, int saveToFile) {
    ULONGLONG elapsedSec;
    LONG peak;
    LONG binPeak;
    double pct;
    uint32_t dataBytes;
    int h, m, s;
    const char* statusText;
    char bytesText[32];
    double dataMB;

    EnterCriticalSection(&g_app.lock);
    peak = g_app.lastPeak;
    binPeak = g_app.currentBinPeak;
    dataBytes = g_app.dataBytesWritten;
    LeaveCriticalSection(&g_app.lock);

    if (binPeak > peak) peak = binPeak;
    pct = (peak / 32767.0) * 100.0;
    if (pct < 0.0) pct = 0.0;
    if (pct > 100.0) pct = 100.0;

    elapsedSec = (GetTickCount64() - g_app.startTickMs) / 1000ULL;
    seconds_to_hms(elapsedSec, &h, &m, &s);

    if (pct < 20.0) statusText = "Volume is low";
    else if (pct > 80.0) statusText = "Volume is too high";
    else statusText = "Volume level is good";

    format_uint_with_commas(dataBytes, bytesText, sizeof(bytesText));
    dataMB = (double)dataBytes / (1024.0 * 1024.0);

    move_cursor_home();
    printf("CasinoLove Voice Recorder 1.0\n");
    printf("============================================================\n");
    printf("Mode:      %s\n", saveToFile ? "Recording" : "Microphone test");
    printf("Input:     %s\n", g_app.inputDeviceName);
    printf("Monitor:   %s\n", g_app.settings.monitorEnabled ? g_app.outputDeviceName : "OFF");
    printf("Format:    %lu Hz | %s | 16-bit PCM\n", (unsigned long)g_app.settings.sampleRate, channels_name(g_app.settings.channels));
    printf("Refresh:   %lu ms\n", (unsigned long)g_app.settings.refreshIntervalMs);
    if (saveToFile) printf("File:      %s\n", g_app.settings.outputPath);
    else printf("File:      (not recording)\n");
    printf("Elapsed:   %02d:%02d:%02d\n", h, m, s);
    printf("Peak:      %5ld / 32767  (%6.2f%%)\n", peak, pct);
    printf("Status:    %-32s\n", statusText);
    if (saveToFile) printf("Data:      %s bytes (%.2f MB)\n", bytesText, dataMB);
    else printf("Data:      %-32s\n", "live test only");
    printf("Stop:      Press ENTER or Ctrl+C\n");
    printf("\nPeak history by update interval, oldest on the left, newest on the right\n");
    render_history_chart(history, historyCount);
    printf("\n");
}

static void flush_keyboard_buffer(void) {
    while (_kbhit()) _getch();
}

static int run_audio_session(const AppSettings* settings, int saveToFile) {
    MMRESULT mmr;
    WAVEHDR headers[INPUT_BUFFER_COUNT];
    BYTE* buffers[INPUT_BUFFER_COUNT] = {0};
    HistoryEntry history[HISTORY_COLS];
    int historyCount = 0;
    ULONGLONG lastCommittedSlot = 0;
    ULONGLONG lastFileSyncMs = 0;
    int i;

    memset(&g_app, 0, sizeof(g_app));
    memset(headers, 0, sizeof(headers));
    memset(history, 0, sizeof(history));
    memset(g_outputBlocks, 0, sizeof(g_outputBlocks));
    g_outputCount = 0;

    g_app.settings = *settings;
    get_input_device_name(g_app.settings.inputUseMapper, g_app.settings.inputDeviceId, g_app.inputDeviceName, sizeof(g_app.inputDeviceName));
    get_output_device_name(g_app.settings.outputUseMapper, g_app.settings.outputDeviceId, g_app.outputDeviceName, sizeof(g_app.outputDeviceName));

    InitializeCriticalSection(&g_app.lock);

    build_format(g_app.settings.sampleRate, g_app.settings.channels, &g_app.fmt);

    if (saveToFile) {
        g_app.outFile = fopen(g_app.settings.outputPath, "wb");
        if (!g_app.outFile) {
            printf("Failed to open output file: %s\n", g_app.settings.outputPath);
            DeleteCriticalSection(&g_app.lock);
            return 1;
        }

        setvbuf(g_app.outFile, NULL, _IOFBF, 262144);

        if (!write_placeholder_header(g_app.outFile, &g_app.fmt)) {
            printf("Failed to write WAV header.\n");
            fclose(g_app.outFile);
            DeleteCriticalSection(&g_app.lock);
            return 1;
        }

        sync_recording_file_locked();
    }

    mmr = waveInOpen(
        &g_app.hWaveIn,
        g_app.settings.inputUseMapper ? WAVE_MAPPER : g_app.settings.inputDeviceId,
        &g_app.fmt,
        (DWORD_PTR)wave_in_proc,
        0,
        CALLBACK_FUNCTION
    );
    if (mmr != MMSYSERR_NOERROR) {
        printf("waveInOpen failed. The selected device may not support this format.\n");
        if (g_app.outFile) fclose(g_app.outFile);
        DeleteCriticalSection(&g_app.lock);
        return 1;
    }

    if (g_app.settings.monitorEnabled) {
        mmr = waveOutOpen(
            &g_app.hWaveOut,
            g_app.settings.outputUseMapper ? WAVE_MAPPER : g_app.settings.outputDeviceId,
            &g_app.fmt,
            0,
            0,
            CALLBACK_NULL
        );
        if (mmr != MMSYSERR_NOERROR) {
            printf("waveOutOpen failed for live monitoring.\n");
            waveInClose(g_app.hWaveIn);
            if (g_app.outFile) fclose(g_app.outFile);
            DeleteCriticalSection(&g_app.lock);
            return 1;
        }
    }

    for (i = 0; i < INPUT_BUFFER_COUNT; ++i) {
        buffers[i] = (BYTE*)malloc(INPUT_BUFFER_BYTES);
        if (!buffers[i]) {
            printf("Failed to allocate input buffer.\n");
            InterlockedExchange(&g_app.running, 0);
            goto cleanup;
        }

        headers[i].lpData = (LPSTR)buffers[i];
        headers[i].dwBufferLength = INPUT_BUFFER_BYTES;

        mmr = waveInPrepareHeader(g_app.hWaveIn, &headers[i], sizeof(headers[i]));
        if (mmr != MMSYSERR_NOERROR) {
            printf("waveInPrepareHeader failed.\n");
            InterlockedExchange(&g_app.running, 0);
            goto cleanup;
        }

        mmr = waveInAddBuffer(g_app.hWaveIn, &headers[i], sizeof(headers[i]));
        if (mmr != MMSYSERR_NOERROR) {
            printf("waveInAddBuffer failed.\n");
            InterlockedExchange(&g_app.running, 0);
            goto cleanup;
        }
    }

    SetConsoleCtrlHandler(console_ctrl_handler, TRUE);
    flush_keyboard_buffer();
    InterlockedExchange(&g_app.running, 1);
    g_app.lastPeak = 0;
    g_app.currentBinPeak = 0;
    g_app.startTickMs = GetTickCount64();

    mmr = waveInStart(g_app.hWaveIn);
    if (mmr != MMSYSERR_NOERROR) {
        printf("waveInStart failed.\n");
        InterlockedExchange(&g_app.running, 0);
        goto cleanup;
    }

    clear_console_screen();

    while (InterlockedCompareExchange(&g_app.running, 0, 0)) {
        ULONGLONG elapsedMs = GetTickCount64() - g_app.startTickMs;
        ULONGLONG elapsedSec = elapsedMs / 1000ULL;
        ULONGLONG elapsedSlots = elapsedMs / (ULONGLONG)g_app.settings.refreshIntervalMs;

        while (lastCommittedSlot < elapsedSlots) {
            LONG slotPeak;
            EnterCriticalSection(&g_app.lock);
            slotPeak = g_app.currentBinPeak;
            g_app.currentBinPeak = 0;
            LeaveCriticalSection(&g_app.lock);

            if (historyCount < HISTORY_COLS) {
                history[historyCount].slotIndex = lastCommittedSlot;
                history[historyCount].peak = slotPeak;
                historyCount++;
            } else {
                int j;
                for (j = 1; j < HISTORY_COLS; ++j) history[j - 1] = history[j];
                history[HISTORY_COLS - 1].slotIndex = lastCommittedSlot;
                history[HISTORY_COLS - 1].peak = slotPeak;
            }
            lastCommittedSlot++;
        }

        if (saveToFile && (elapsedMs - lastFileSyncMs) >= FILE_SYNC_INTERVAL_MS) {
            EnterCriticalSection(&g_app.lock);
            sync_recording_file_locked();
            LeaveCriticalSection(&g_app.lock);
            lastFileSyncMs = elapsedMs;
        }

        (void)elapsedSec;
        render_session_screen(history, historyCount, saveToFile);

        if (_kbhit()) {
            int ch = _getch();
            if (ch == 13) {
                InterlockedExchange(&g_app.running, 0);
                break;
            }
        }

        Sleep(g_app.settings.refreshIntervalMs);
    }

cleanup:
    if (g_app.hWaveIn) {
        waveInStop(g_app.hWaveIn);
        waveInReset(g_app.hWaveIn);
    }

    Sleep(200);

    for (i = 0; i < INPUT_BUFFER_COUNT; ++i) {
        if (g_app.hWaveIn && headers[i].lpData) {
            waveInUnprepareHeader(g_app.hWaveIn, &headers[i], sizeof(headers[i]));
        }
        free(buffers[i]);
        buffers[i] = NULL;
    }

    if (g_app.hWaveIn) {
        waveInClose(g_app.hWaveIn);
        g_app.hWaveIn = NULL;
    }

    if (g_app.hWaveOut) {
        waveOutReset(g_app.hWaveOut);
        cleanup_finished_output_blocks();
        for (i = 0; i < MAX_OUTPUT_QUEUE; ++i) {
            OutputBlock* block = g_outputBlocks[i];
            if (block) {
                while (!(block->hdr.dwFlags & WHDR_DONE)) Sleep(1);
                waveOutUnprepareHeader(g_app.hWaveOut, &block->hdr, sizeof(block->hdr));
                free(block->data);
                free(block);
                g_outputBlocks[i] = NULL;
            }
        }
        waveOutClose(g_app.hWaveOut);
        g_app.hWaveOut = NULL;
    }

    if (g_app.outFile) {
        EnterCriticalSection(&g_app.lock);
        sync_recording_file_locked();
        LeaveCriticalSection(&g_app.lock);
        fclose(g_app.outFile);
        g_app.outFile = NULL;
    }

    DeleteCriticalSection(&g_app.lock);
    SetConsoleCtrlHandler(console_ctrl_handler, FALSE);
    flush_keyboard_buffer();

    return 0;
}

static void choose_output_filename(AppSettings* settings) {
    char line[MAX_PATH];
    printf("Current output file: %s\n", settings->outputPath);
    read_line("Enter new output WAV filename: ", line, sizeof(line));
    if (line[0] != '\0') snprintf(settings->outputPath, sizeof(settings->outputPath), "%s", line);
}

static void choose_input_device(AppSettings* settings) {
    UINT count = waveInGetNumDevs();
    int choice;
    UINT i;

    clear_console_screen();
    printf("Select input device\n\n");
    printf("0) Default system input\n");
    for (i = 0; i < count; ++i) {
        WAVEINCAPSA caps;
        if (waveInGetDevCapsA(i, &caps, sizeof(caps)) == MMSYSERR_NOERROR) printf("%u) %s\n", i + 1, caps.szPname);
        else printf("%u) Input device %u\n", i + 1, i);
    }
    printf("\n");
    if (!read_menu_int("Choice: ", 0, (int)count, &choice)) {
        printf("Invalid choice.\n");
        wait_for_enter();
        return;
    }

    if (choice == 0) {
        settings->inputUseMapper = 1;
        settings->inputDeviceId = 0;
    } else {
        settings->inputUseMapper = 0;
        settings->inputDeviceId = (UINT)(choice - 1);
    }
    choose_best_defaults(settings);
}

static void choose_output_device(AppSettings* settings) {
    UINT count = waveOutGetNumDevs();
    int choice;
    UINT i;

    clear_console_screen();
    printf("Select monitoring output device\n\n");
    printf("0) Default system output\n");
    for (i = 0; i < count; ++i) {
        WAVEOUTCAPSA caps;
        if (waveOutGetDevCapsA(i, &caps, sizeof(caps)) == MMSYSERR_NOERROR) printf("%u) %s\n", i + 1, caps.szPname);
        else printf("%u) Output device %u\n", i + 1, i);
    }
    printf("\n");
    if (!read_menu_int("Choice: ", 0, (int)count, &choice)) {
        printf("Invalid choice.\n");
        wait_for_enter();
        return;
    }

    if (choice == 0) {
        settings->outputUseMapper = 1;
        settings->outputDeviceId = 0;
    } else {
        settings->outputUseMapper = 0;
        settings->outputDeviceId = (UINT)(choice - 1);
    }
    adjust_to_supported_combo(settings);
}

static void toggle_monitoring(AppSettings* settings) {
    settings->monitorEnabled = !settings->monitorEnabled;
    adjust_to_supported_combo(settings);
}

static void choose_sample_rate(AppSettings* settings) {
    int choice;
    int i;

    clear_console_screen();
    printf("Select sample rate\n\n");
    for (i = 0; i < 3; ++i) {
        int supported = combo_supported(settings, kRates[i], settings->channels);
        printf("%d) %6lu Hz   %s\n", i + 1, (unsigned long)kRates[i], supported ? "supported" : "not supported");
    }
    printf("\nCurrent channels: %s\n\n", channels_name(settings->channels));

    if (!read_menu_int("Choice: ", 1, 3, &choice)) {
        printf("Invalid choice.\n");
        wait_for_enter();
        return;
    }

    if (!combo_supported(settings, kRates[choice - 1], settings->channels)) {
        printf("That sample rate is not supported for the current settings.\n");
        wait_for_enter();
        return;
    }

    settings->sampleRate = kRates[choice - 1];
}

static void choose_channels(AppSettings* settings) {
    int choice;

    clear_console_screen();
    printf("Select channels\n\n");
    printf("1) Mono    %s\n", combo_supported(settings, settings->sampleRate, 1) ? "supported" : "not supported");
    printf("2) Stereo  %s\n", combo_supported(settings, settings->sampleRate, 2) ? "supported" : "not supported");
    printf("\nCurrent sample rate: %lu Hz\n\n", (unsigned long)settings->sampleRate);

    if (!read_menu_int("Choice: ", 1, 2, &choice)) {
        printf("Invalid choice.\n");
        wait_for_enter();
        return;
    }

    if (!combo_supported(settings, settings->sampleRate, (WORD)choice)) {
        printf("That channel mode is not supported for the current settings.\n");
        wait_for_enter();
        return;
    }

    settings->channels = (WORD)choice;
}

static void choose_recording_format_notice(void) {
    clear_console_screen();
    printf("Recording format\n\n");
    printf("This version only supports 16-bit recording.\n");
    wait_for_enter();
}

static void choose_refresh_interval(AppSettings* settings) {
    int choice;
    int i;

    clear_console_screen();
    printf("Select meter and graph update interval\n\n");
    for (i = 0; i < 5; ++i) {
        printf("%d) %lu ms\n", i + 1, (unsigned long)kRefreshChoices[i]);
    }
    printf("\nCurrent refresh interval: %lu ms\n\n", (unsigned long)settings->refreshIntervalMs);

    if (!read_menu_int("Choice: ", 1, 5, &choice)) {
        printf("Invalid choice.\n");
        wait_for_enter();
        return;
    }

    settings->refreshIntervalMs = kRefreshChoices[choice - 1];
}

static void print_main_menu(const AppSettings* settings) {
    char inputName[256];
    char outputName[256];

    get_input_device_name(settings->inputUseMapper, settings->inputDeviceId, inputName, sizeof(inputName));
    get_output_device_name(settings->outputUseMapper, settings->outputDeviceId, outputName, sizeof(outputName));

    clear_console_screen();
    printf("CasinoLove Voice Recorder 1.0\n");
    printf("============================================================\n\n");
    printf("Settings\n");
    printf("--------\n");
    printf("1) Output filename         : %s\n", settings->outputPath);
    printf("2) Input device            : %s\n", inputName);
    printf("3) Monitoring output device: %s\n", outputName);
    printf("4) Live playback sound     : %s\n", settings->monitorEnabled ? "ON" : "OFF");
    printf("5) Sample rate             : %lu Hz\n", (unsigned long)settings->sampleRate);
    printf("6) Channels                : %s\n", channels_name(settings->channels));
    printf("7) Recording format        : 16-bit PCM\n");
    printf("8) Meter update interval   : %lu ms\n", (unsigned long)settings->refreshIntervalMs);
    printf("\nActions\n");
    printf("-------\n");
    printf("9) Microphone test\n");
    printf("10) Start recording\n");
    printf("0) Exit\n");
}

static void run_menu_loop(void) {
    AppSettings settings;
    int running = 1;
    int choice;

    memset(&settings, 0, sizeof(settings));
    snprintf(settings.outputPath, sizeof(settings.outputPath), "recording.wav");
    settings.inputUseMapper = 1;
    settings.inputDeviceId = 0;
    settings.outputUseMapper = 1;
    settings.outputDeviceId = 0;
    settings.monitorEnabled = 0;
    settings.sampleRate = 48000;
    settings.channels = 1;
    settings.refreshIntervalMs = 50;

    choose_best_defaults(&settings);

    while (running) {
        print_main_menu(&settings);

        if (!read_menu_int("\nChoice: ", 0, 10, &choice)) {
            printf("Invalid choice.\n");
            wait_for_enter();
            continue;
        }

        switch (choice) {
        case 1:
            choose_output_filename(&settings);
            break;
        case 2:
            choose_input_device(&settings);
            break;
        case 3:
            choose_output_device(&settings);
            break;
        case 4:
            toggle_monitoring(&settings);
            break;
        case 5:
            choose_sample_rate(&settings);
            break;
        case 6:
            choose_channels(&settings);
            break;
        case 7:
            choose_recording_format_notice();
            break;
        case 8:
            choose_refresh_interval(&settings);
            break;
        case 9:
            if (!current_combo_supported(&settings)) {
                printf("\nThe current input/output/format combination is not supported.\n");
                wait_for_enter();
                break;
            }
            run_audio_session(&settings, 0);
            printf("\nMicrophone test finished.\n");
            wait_for_enter();
            break;
        case 10:
            if (!current_combo_supported(&settings)) {
                printf("\nThe current input/output/format combination is not supported.\n");
                wait_for_enter();
                break;
            }
            if (file_exists(settings.outputPath) && !ask_yes_no("\nThat file already exists. Overwrite it? [y/N]: ", 0)) {
                break;
            }
            run_audio_session(&settings, 1);
            printf("\nRecording finished.\n");
            wait_for_enter();
            break;
        case 0:
            running = 0;
            break;
        default:
            break;
        }
    }
}

int main(void) {
    run_menu_loop();
    return 0;
}
