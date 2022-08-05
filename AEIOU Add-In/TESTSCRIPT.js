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

// // one of the ways to pass parameters into the py file if the other one doesnt work
// script.stdin.write(data);
// // End data write
// script.stdin.end();

script.on('close', (code) => {
    console.log("Process quit with code : " + code);
    // res.send(datatoSend);
});

// node TESTCRIPT.js