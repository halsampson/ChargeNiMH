// ChargeNiMH.cpp 

// Charge cycle NiMH cell

// Using Agilent E3631A
//   P25V: charge
//   N25V: discharge thru LM7805 "Zener" (Gnd, Out shorted)
//   P6V:  read 4 wire voltage

// cycles NiMH cell
// Data: https://docs.google.com/spreadsheets/d/1tSRwlEcyB1IPcc4s9cZxGvZLJLEmf76hiC5O18wYq60/edit?usp=sharing

// TODO: model R1 - C1 - R2 - C2 ...  --> capacity estimate 
//   sweep pulse length
//   sweep also pulse height
//   analyze responses
//   5 minute tau
//   balance charge/discharge for stable SoC
//   also estimate internal discharge (slow)

// TODO: add temperature sensor

// TODO: Try very short, high current pulsed charge/discharge cycles to break up crystallization
//   Any way to remove surface oxide layers?

#include <windows.h>
#include <time.h>
#include <stdio.h>
#include <conio.h>
#include <math.h>

const float iMax = 1.0; // Agilent N25V, P25V
const int MaxReadbackMillisec = 100;

#define COM_PORT "COM2"

HANDLE hCom;

void openSerial() {
	hCom = CreateFileA("\\\\.\\" COM_PORT, GENERIC_WRITE | GENERIC_READ, 0, NULL, OPEN_EXISTING, NULL, NULL);
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

float vMax;

float vInternal, vExternal;
float isr;
float iBump = 0.1;  // < min discharge current

void batteryISR() {
	float vComply = vMax + iBatt * 0.2; // TODO: meter, cable + any diode drop from P25V
	setVI(vComply, iBatt);  
	float vExternal1 = getV();

	if (iBump > fabs(iBatt)) iBump = fabs(iBatt);
	const float MaxISR = 5; // Ohms
  setVI(vComply + MaxISR * iBump, iBatt += iBump);
	float vBump = getV();

	setVI(vComply, iBatt -= iBump);
	vExternal = getV();

	// avg pop/dip to avoid charging delta V error
	isr = (vBump - (vExternal1 + vExternal) / 2) / iBump;
	vInternal = vExternal - isr * iBatt;
}

int displayOnSecs = 15;  // initial

float vTerminateDischarge;
float mAh, mWh;

bool terminate() { 
	vExternal = getV();
	vInternal = vExternal - isr * iBatt;
	if (iBatt < 0) {
    if (vInternal <= vTerminateDischarge) return true;  // terminate
  } else {
    if (vExternal >= vMax) return true;  // terminate	TODO: batter vInternal? (vs. heating)
  }
	return false;
}

bool toggleCharging;
float prevVext;

bool report(int reportMinutes) {
	ULONGLONG msStart = GetTickCount64();
	toggleCharging = false;

	batteryISR();
	printf("%.4f,%5.1f,%4.0f,%7.4f,%5.0f,%5.0f\n", vExternal, (vExternal - prevVext) * 1000, isr * 1000, vInternal, mAh, mWh);
	prevVext = vExternal;

	float amps = 0;
  ULONGLONG msEnd = msStart + reportMinutes * 60 * 1000;
	while (1) {
		if (terminate()) return true;

		amps = getI(); 
		if (fabs(amps) < 0.004) {
			printf("Open\n");
			return false;  // open circuit; 4mA max offset
		}
		
    if (_kbhit()) // beware Escaped chars!!
			switch (_getch()) {
        case 't' : toggleCharging = true; return false;  // toggleCharging charge/discharge
        default : displayOnSecs = _getch() - '0'; break;
      }

		if (displayOnSecs > 0) {
			// VFD filament is direct from xfrmr;  122 vs. 115 VAC -> < 12.5% extra filament power
			char display[12 + 12 + 1 + 7]; // 12 chars not counting .,;
			sprintf_s(display, sizeof display, "DISP:TEXT \"%.4fV %.0fmO\"", vExternal, isr * 1000);
			cmd(display);
		} else if (!displayOnSecs) 
			cmd("DISP Off");

		Sleep(1000);
		--displayOnSecs;

		ULONGLONG msNow = GetTickCount64();
		if (msNow >= msEnd) {	
			float delta_mAh = amps * (msNow - msStart) / 3600;
			mAh += delta_mAh;
			mWh += vInternal * delta_mAh; 
			return true;
		}
	}
}


float C; // Cell Capacity in Ah

bool charge() {
	mAh = 0, mWh = 0;

	while (vInternal < 1) { // initial slow charge
	  iBatt = C/10;
		report(1);
	}

#if 1
	iBatt = min(iMax - iBump, C); // dV bump of ~8 mV at C/4
#endif	

	int levelMins = 0;
	float vPeak = 0;
	float avgISR = 0;

	while (1) {
		const int reportMinutes = 1;   // for plot
		if (!report(reportMinutes)) return false;		
		
		// see https://www.powerstream.com/NiMH.htm  for charge termination ideas

		if (vExternal >= vMax) break; // terminate
		if (vInternal >= 1.6) break;  // TODO: vInternal still depends on iBatt due to complex model; adjust ****

		// TODO: terminate based on previous measured discharge mAh??
		if (mAh >= C * 1000 * 1.1) break; // + 40% below: 150% typ. required to fill good cell at fast charge

		// TODO: better detect 2nd vExternal downward inflection
		// best dV signal from vExternal  https://lygte-info.dk/info/batteryChargingNiMH%20UK.html
 		if (vExternal >= vPeak + 0.0005) {  // terminate on very slow rise as well as negative dVext
			vPeak = vExternal;
			levelMins = 0;
		}	else if (vExternal > vPeak)
			vPeak = vExternal;
		else if (vInternal > 1.45 && mAh > C * 1000 * 0.7) { // terminate only after 2nd peak in dVexternal
			displayOnSecs = reportMinutes * 60; // to watch termination
			if (vExternal <= vPeak - 0.001)  
			  break; // terminate
		  else if ((levelMins += reportMinutes) >= 20)  // in case dv/dt not seen at low charge rates
			  break; // terminate		  
		}

		// if (isr < avgISR * 0.8) break; // ?ISR can drop quickly as temperature increases near max charge ??
		// avgISR += (isr - avgISR) / 8;
	}
	report(0);

	// C/10 charge rate top off for 4 hours
	if (vExternal < vMax) {
		iBatt = C/10; // safer top off charge rate
		for (int topOff = 4 * 60; (topOff -= 5) >= 0;) { // minutes
			report(5);
			if (mAh >= C * 1000 * 1.2) break;  // TODO: vs. max 50% Coulombic loss
			// TODO: other termination signs?
		}
	}
	report(0);
	return true;
}

bool discharge() {
  mAh = 0, mWh = 0;

	// TODO: try servoing current to get roughly constant dV drop? vs. crystals/oxide/...
	//   would reduce current at both start and end of discharge -- when reactions are concentrated near anode/cathode?

	iBatt = max(-iMax, -C/2); // 1A N25V max 
	vTerminateDischarge = 1.0;   // TODO: or as abs(dVext) or ISR (later) start increasing??
  do {
    if (!report(2)) return false;
  } while (vExternal > vTerminateDischarge);
	report(0);

	iBatt = -C/20; // slow reconditioning discharge to break up crystals
	vTerminateDischarge = 0.4;
	do {
		if (!report(5)) return false;
	} while (vExternal > vTerminateDischarge);
	report(0);
	return true;
}


bool cycleNiMH() {
  C = 3.5; // 0.3; // 3.5;
	vMax = 1.7;  // TODO: depends on battery age, temperature - better vInternalMax ****

  if (!discharge() && !toggleCharging) return false;
	if (!charge() && !toggleCharging) return false;

	return true;
}


int main() {
	openSerial();

	getResponse(); // flush

	cmd("*IDN?");
	printf("%s", response);

	cmd("*RST");
	cmd("SYST:REM");

	cmd("APPL P6V,4.4,0.002");  // for reading 4-wire voltage
	cmd("OUTP ON");

	// TODO detect battery type, capacity
	while (1) {		
		cmd("DISP:TEXT \"Insert cell\"");
		vMax = 4.3;
		while (1) {
		  if (getV() < vMax) break; // check cell inserted
		  Sleep(1);
		}

		displayOnSecs = 10;
		prevVext = getV();
		while (cycleNiMH());
		setVI(0, 0);
	}
}
