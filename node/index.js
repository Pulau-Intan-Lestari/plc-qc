const Modbus = require('jsmodbus');
const net = require('net');
const socket = require('socket.io')(3000);  // WebSocket server on port 3000
const tcpSocket = new net.Socket();
const client = new Modbus.client.TCP(tcpSocket);

// PLC connection settings
const PLC_IP = '192.168.1.100';  // Replace with your PLC's IP
const PLC_PORT = 502;            // Default Modbus port

// Roller parameters
const rollerCircumference = 40;  // Example: 40 inches per rotation (adjust based on your roller size)
let yardData = 0;

// Function to convert number to binary string
function numToBinaryString(num, length = 16) {
    return num.toString(2).padStart(length, '0');
}

// Function to calculate yardage based on rotation count
function calculateYardage(rotations) {
    const distanceInInches = rotations * rollerCircumference;
    return distanceInInches / 36;  // Convert inches to yards
}

// Function to handle and parse PLC data
function handlePLCData(plcData) {
    console.log("PLC Data:", plcData);

    // Assuming register 5 or 6 contains rotation data
    const rotations = plcData[5];  // Example: Assuming plcData[5] tracks roller rotations

    // Calculate yardage
    yardData = calculateYardage(rotations);
    console.log("Calculated Yardage:", yardData);
    
    // Send data to frontend via WebSocket
    socket.emit('yardData', yardData);
}

// Connect to the PLC and read data
tcpSocket.on('connect', function () {
    console.log('Connected to PLC!');

    // Polling PLC for data every 1 second (adjust interval as needed)
    setInterval(() => {
        client.readHoldingRegisters(0, 10)
            .then(function (response) {
                const plcData = response.response._body.values;
                handlePLCData(plcData);  // Process and send the PLC data
            })
            .catch(console.error);
    }, 1000);  // 1000 ms = 1 second interval
});

// Handle socket connection errors
tcpSocket.on('error', function (err) {
    console.error('Connection error:', err);
});

// Handle PLC disconnection
tcpSocket.on('close', function () {
    console.log('Disconnected from PLC, attempting to reconnect...');
    // Attempt to reconnect every 5 seconds
    setTimeout(() => tcpSocket.connect({ host: PLC_IP, port: PLC_PORT }), 5000);
});

// Handle PLC reconnection
tcpSocket.on('timeout', function () {
    console.error('PLC connection timed out, reconnecting...');
    tcpSocket.connect({ host: PLC_IP, port: PLC_PORT });
});

// Connect to the PLC
tcpSocket.connect({ host: PLC_IP, port: PLC_PORT });

// WebSocket connection to frontend
socket.on('connection', (clientSocket) => {
    console.log('Frontend connected');
    // Send initial yard data when frontend connects
    clientSocket.emit('yardData', yardData);
});
