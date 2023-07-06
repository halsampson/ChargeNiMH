// AgBatteryISR.cpp 

// Using Agilent E3631A
//   P25V: charge
//   N25V: discharge thru LM7805 "Zener" (Gnd, Out shorted)
//   P6V:  read 4 wire voltage

// cycles NiMH cell
// Data: https://docs.google.com/spreadsheets/d/1tSRwlEcyB1IPcc4s9cZxGvZLJLEmf76hiC5O18wYq60/edit?usp=sharing


#include <windows.h>
#include <time.h>
#include <stdio.h>
#include <conio.h>
#include <math.h>

const int MaxReadbackMillisec = 100;
// longer -> less recovery slope

HANDLE hCom;

void openSerial() {
	hCom = CreateFileA("\\\\.\\COM2", GENERIC_WRITE | GENERIC_READ, 0, NULL, OPEN_EXISTING, NULL, NULL);
	if (hCom == INVALID_HANDLE_VALUE) exit(-2);

	DCB dcb = { 0 };
	dcb.DCBlength = sizeof(DCB);
	dcb.BaudRate = 9600;
	dcb.ByteSize = 8;
	dcb.StopBits = 2;
	dcb.fParity = FALSE;
	dcb.fBinary = TRUE;
	dcb.fOutxDsrFlow = true;
	dcb.fDtrControl = DTR_CONTROL_HANDSHAKE;
	dcb.fRtsControl =  RTS_CONTROL_ENABLE;

	if (!SetCommState(hCom, &dcb)) exit(-3);

	COMMTIMEOUTS timeouts = { 0 };  // in ms
	timeouts.ReadIntervalTimeout = 50; // between characters
	timeouts.ReadTotalTimeoutMultiplier = 2; // * num requested chars
	timeouts.ReadTotalTimeoutConstant = MaxReadbackMillisec + 100; // + this = total timeout    100 ms max readback delay
	if (!SetCommTimeouts(hCom, &timeouts)) exit(-4);
}


int rxRdy() {
	COMSTAT cs;
	DWORD commErrors;
	if (!ClearCommError(hCom, &commErrors, &cs)) return -1;
	return cs.cbInQue;
}

char response[40];  // IDN length

void getResponse() {
	DWORD bytesRead = 0;
	if (!ReadFile(hCom, response, sizeof response - 1, &bytesRead, NULL))
		bytesRead = 0;
	response[bytesRead] = 0;
	if (!bytesRead) printf("?");
}

void cmd(const char* cmd) {
	if (!hCom) openSerial();
	WriteFile(hCom, cmd, strlen(cmd), NULL, NULL);
	WriteFile(hCom, "\n", 1, NULL, NULL);
	if (strchr(cmd, '?')) {
		getResponse();
		if (!response[0]) printf("** No response to: %s\n", cmd);
	}
}

void setVI(float volts, float amps) {
	char cmdStr[64];

	if (amps >= 0)
		sprintf_s(cmdStr, sizeof cmdStr, "APPL P25V,%.3f,%.3f;APPL N25V,0,0", volts, amps);
	else sprintf_s(cmdStr, sizeof cmdStr, "APPL P25V,%.3f,0;APPL N25V,-4.6,%.3f", volts, -amps);  // LM7805 4.5V @ 1A
	cmd(cmdStr);
}

float getV() {
	cmd("MEAS:VOLT? P6V");
	return atof(response);
}

float iBatt;

float getI() {
	if (iBatt > 0) {
		cmd("MEAS:CURR? P25V");
	  return atof(response);
	} else {
		cmd("MEAS:CURR? N25V");   // positive
    return -atof(response);
	}
}

float vMax, vComply;

float vInternal, vExternal2;
float isr;

void batteryISR() {
	const float iBump = 0.1;  // < min discharge current

	vComply = vMax + iBatt * 0.2; // TODO: meter, cable + any diode drop from P25V
	setVI(vComply, iBatt);  
	float vExternal1 = getV();

	const float MaxISR = 5; // Ohms
  setVI(vComply + MaxISR * iBump, iBatt += iBump);
	float vBump = getV();

	setVI(vComply, iBatt -= iBump);
	vExternal2 = getV();

	// avg pop/dip to avoid charging delta V error
	isr = (vBump - (vExternal1 + vExternal2) / 2) / iBump;
	vInternal = vExternal2 - isr * iBatt;
}

int toggle;

int displayOnSecs = 15;  // initial

// TODO: model R1 - C1 - R2 - C2  --> capacity estimate 
float vTerminate;
float vExternal;
float mAh, mWh;

bool terminate() { 
	vExternal = getV();
	if (iBatt < 0) {
    if (vExternal <= vTerminate) return true;  // terminate	
  } else {
    if (vExternal >= vComply) return true;  // terminate	
  }
	return false;
}

bool report(int reportMinutes) {
	ULONGLONG msStart = GetTickCount64();

	if (terminate()) return true;

	batteryISR();
	printf("%.4f,%.3f,%.4f,%5.0f,%5.0f\n", vExternal2, isr, vInternal, mAh, mWh);

	float amps;
  ULONGLONG msEnd = msStart + reportMinutes * 60 * 1000;
	while (GetTickCount64() < msEnd) {
		if (terminate()) return true;

		amps = getI(); 
		if (fabs(amps) < 0.01) {
			printf("Open\n");
			return false;  // open circuit; 4mA max offset
		}

		// filament is direct from xfrmr;  122 vs. 115 VAC -> < 12.5% extra filament power
    if (_kbhit()) 
			switch (_getch()) {
        case ' ' : toggle = 1; return false;  // TODO: toggle charge/discharge
        default : displayOnSecs = _getch() - '0'; break;
      }

		if (displayOnSecs > 0) {
			char display[12 + 12 + 1 + 7]; // 12 chars not counting .,;
			sprintf_s(display, sizeof display, "DISP:TEXT \"%.4fV %.0fmO\"", vExternal, isr * 1000);
			cmd(display);
		} else if (!displayOnSecs) 
			cmd("DISP Off");
	  --displayOnSecs;
		Sleep(1000);
	}

	float delta_mAh = amps * (GetTickCount64() - msStart) / 3600;
	mAh += delta_mAh;
	mWh += vInternal * delta_mAh; 
	return true;
}


float C; // Capacity in Ah

bool charge() {
#if 1
	iBatt = C/4; // dV bump of ~8 mV at C/4
	vMax = 1.6;  // vMax < 1.57 for GOOD cell + ISR
#else
	iBatt = C/10; // forming charge
	vMax = 1.53;
#endif

	mAh = 0, mWh = 0;
	int levelMins = 0;
	float vPeak = 0;
	float avgISR = 0;
	while (1) {
		const int reportMinutes = 1;   // for plot
		if (!report(reportMinutes)) return false;		
		
		if (vExternal >= vComply) break; // terminate	  TODO: high?
		
		// if (isr < avgISR * 0.8) break; // ISR can drop quickly as temperature increases near max charge ??
		// avgISR += (isr - avgISR) / 8;

		// TODO: find end of 2nd downward inflection
		// V slope high at first, levels off, near end slope increases again, levels off
		//       2nd derivative: negative    0               positive      negative    0

		// better dV signal from vExternal  https://lygte-info.dk/info/batteryChargingNiMH%20UK.html
 		if (vExternal >= vPeak + 0.0005) {  // terminate on very slow rise as well as negative dVext
			vPeak = vExternal;
			levelMins = 0;
		}	else if (vExternal > vPeak)
			vPeak = vExternal;
		else if (vExternal <= vPeak - 0.001)  // detect any bump - TODO: beware early charge ISR drop, noise
			break; // terminate

		if ((levelMins += reportMinutes) >= 20)  // in case dv/dt not seen at low charge rates
			break; // terminate
		else displayOnSecs = reportMinutes * 60; // to watch termination
	}
	report(0);
	return true;
}

bool discharge() {
  mAh = 0, mWh = 0;

	vTerminate = 1.0;
	iBatt = -1.0; // Agilent N25V max
  do {
    if (!report(2)) return false;
  } while (vExternal > vTerminate); 

	vTerminate = 0.4;
	iBatt = -C/20; // reconditioning discharge to break up crystals
	do {
		if (!report(5)) return false;
	} while (vExternal > vTerminate); // NiCd deep discharge
	report(0);
	return true;
}


bool cycleNiMH() {
  C = 3.5;
	vMax = 1.6;  // TODO: depends on temperature

  if (!discharge() && --toggle) return false;
	if (!charge() && --toggle) return false;

	return true;
}


int main() {
	openSerial();

	getResponse(); // flush

	cmd("*IDN?");
	printf("%s", response); // HEWLETT-PACKARD,E3631A,0,2.1-5.0-1.0\r\n

	cmd("*RST");
	cmd("SYST:REM");

	cmd("DISP:TEXT \"Insert cell\"");

	cmd("APPL P6V,4.4,0.002");  // for reading 4-wire voltage
	cmd("OUTP ON");

	// TODO detect battery type
	while (1) {
		while (cycleNiMH());
		setVI(0, 0);
		// TODO: check for cell inserted
		_getch();
	}
}
