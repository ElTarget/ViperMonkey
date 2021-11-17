# Overview

This dockerfile takes the base ubuntu docker image and layers on
software necessary to run ViperMonkey. 

Note for running docker commands: the following commands use sudo, but
you can also add your user to the docker group to avoid this.

## Using This Dockerfile

### Step 1 - build the docker image
Build it!  'sudo docker build -t vipermonkey:latest .'

NOTE: There is a period at the end of the command.  This tells docker
to look for the dockerfile in the current directory.  It is expected
to take 5+ minutes to build.

### Step 2 - push to Docker Hub

Log into your Docker hub account.

```
sudo docker login
```

Push the container to Docker Hub.

```
sudo docker push kirksayre/vipermonkey:latest
```

