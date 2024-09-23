const Modbus = require('jsmodbus');
const net = require('net');
const socket = new net.Socket();
const client = new Modbus.client.TCP(socket);

// PLC connection details (Update these)
const PLC_IP = '192.168.3.4';  // Replace with your PLC's IP address
const PLC_PORT = 502;            // Default Modbus port

// Function to convert a number to a binary string
function numToBinaryString(num, length = 16) {
    return num.toString(2).padStart(length, '0');
}

// Function to parse PLC data and log binary and meaningful states
function parsePLCData(plcData) {
    console.log("Raw PLC Data:", plcData);

    // Convert PLC data (numeric) into binary strings
    const binX0 = numToBinaryString(plcData[5]);
    const binX20 = numToBinaryString(plcData[6]);
    const binY0 = numToBinaryString(plcData[7]);
    const binY20 = numToBinaryString(plcData[8]);

    console.log("Binary X0 (Input Signals):", binX0);
    console.log("Binary X20 (Input Signals):", binX20);
    console.log("Binary Y0 (Output Signals):", binY0);
    console.log("Binary Y20 (Output Signals):", binY20);

    // Parse the binary strings into meaningful values (e.g., Machine ON/OFF)
    const machineOn = binY0[0] === '1';
    const forwardMotion = binY0[1] === '1';
    const backwardMotion = binY0[2] === '1';

    console.log(`Machine State: ${machineOn ? 'ON' : 'OFF'}`);
    console.log(`Forward Motion: ${forwardMotion ? 'ON' : 'OFF'}`);
    console.log(`Backward Motion: ${backwardMotion ? 'ON' : 'OFF'}`);
}

// Connect to the PLC and read data
socket.on('connect', function () {
    console.log('Connected to PLC!');

    // Read 10 registers starting from address 0
    client.readHoldingRegisters(0, 10)
        .then(function (response) {
            const plcData = response.response._body.values;
            parsePLCData(plcData);  // Pass data to parser
        })
        .catch(console.error)
        .finally(() => socket.end());
});

// Handle socket connection errors
socket.on('error', function (err) {
    console.error('Connection error:', err);
});

// Connect to the PLC
socket.connect({ host: PLC_IP, port: PLC_PORT });
