---
layout: post
section-type: post
title:  "Using Python to stream from a Swann DVR camera system"
date:   2018-04-01
categories: tech
tags: [ 'python', 'home-hacks' ]
---

Here is a small code snippet that you can use to stream from your Swann DVR security system. The DVR resides within the same local network as the *python* code streaming the video.  The code incorporates the username and password in the RTSP call when instantiated.  Take a different approach in the code to make it more secure.  Additionally, Secure RTSP protocol over 443 is also available.  

![videocam](/gallery/videocam.png)
{% highlight ruby %}
#!/usr/bin/env python -c

# Import OpenCV and threading packages
import cv2
import threading

# Define class for the camera thread.
class CamThread(threading.Thread):

    def __init__(self, previewname, camid):
        threading.Thread.__init__(self)
        self.previewname = previewname
        self.camid = camid

    def run(self):
        print("Starting " + self.previewname)
        previewcam(self.previewname, self.camid)

# Function to preview the camera.
def previewcam(previewname, camid):
    cv2.namedWindow(previewname)
    cam = cv2.VideoCapture(camid)
    if cam.isOpened():
        rval, frame = cam.read()
    else:
        rval = False

    while rval:
        cv2.imshow(previewname, frame)
        rval, frame = cam.read()
        key = cv2.waitKey(20)
        if key == 27:  # Press ESC to exit/close each window.
            break
    cv2.destroyWindow(previewname)

# Create different threads for each video stream, then start it.
thread1 = CamThread("Driveway", 'rtsp://username:SuperSecurePassword@192.168.1.2/Streaming/Channels/102')
thread2 = CamThread("Front Door", 'rtsp://username:SuperSecurePassword@192.168.1.2/Streaming/Channels/202')
thread3 = CamThread("Garage", 'rtsp://username:SuperSecurePassword@192.168.1.2/Streaming/Channels/302')
thread1.start()
thread2.start()
thread3.start()

# check out tutorials point link below for multi-threading sample
# https://www.tutorialspoint.com/python/python_multithreading.htm

{% endhighlight %}