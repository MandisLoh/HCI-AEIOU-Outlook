var input = 'your exam results are out'
        
const spawn = require('child_process').spawn;
const script = spawn('py', ['ml copy 2.py', input]);
console.log(input)

script.stdout.on('data', (data) => {
    // datatoSend = data.toString();
    console.log(`${data}`)
});
script.stderr.on('data', (data) => {
    // As said before, convert the Uint8Array to a readable string.
    console.error(`stderr: ${data}`);
});

script.on('close', (code) => {
    console.log("Process quit with code : " + code);
    // res.send(datatoSend);
});

// node TESTCRIPT.js