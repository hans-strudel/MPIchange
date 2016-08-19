pushd "\\btsj0\BTSJ2090\BT-Off\ENG\MPI\Job Card" 
echo %~dp0
node %~dp0\parse.js "%CD%"
pause
popd