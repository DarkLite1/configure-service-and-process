# Goal
The goal of `helloworld.exe` program is to serve as a process that can be killed without impacting any production critical process on the client.

# Create the program
C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\csc.exe .\helloworld.cs

# Run the program
Start-Process -FilePath .\helloworld.exe -NoNewWindow