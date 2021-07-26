FROM ubuntu:focal

# Label the image appropriately
LABEL description="Docker container for running Python 3 ViperMonkey" maintainer="Kirk Sayre"

# Hack to get pypy3 and python3 to install packages in the same place.
RUN mkdir /usr/local/lib/pypy3.7
RUN ln -s /usr/local/lib/pypy3.7 /usr/local/lib/python3.6

# Set environment variables
ENV PATH=/opt/ViperMonkey/vipermonkey:$PATH
ENV PYTHONPATH=/opt/ViperMonkey/vipermonkey

# Set timezone to avoid timezone setup needing interaction.
RUN ln -snf /usr/share/zoneinfo/US /etc/localtime && echo "US" > /etc/timezone

# Run apt-get update and install needed packages
#
# Package list needed for python3-uno taken from https://pypi.org/project/unotools/
RUN apt-get update && apt-get install -y \
        apt-utils \
	git\
	wget\
	build-essential \
        zip\
	file\
	libpython2-dev\
      	libpython3-dev\
	zlib1g-dev\
	python3\
	python3-pip\
	software-properties-common\
        libreoffice\
        libreoffice-script-provider-python\
        libuno-cppu3\
        python3-uno\
        python3\
        unzip\
&& rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /opt/

# Python 3 specific packages used by export_all_sheets.py.
RUN pip3 install psutil
RUN pip3 install unotools
RUN pip3 install unoconv

# Install pypy3

# Base pypy3 install.
RUN add-apt-repository -y ppa:pypy/ppa
RUN apt update
RUN apt install -y pypy3

# Install pypy3 pip.
RUN wget https://bootstrap.pypa.io/get-pip.py
RUN pypy3 get-pip.py

# Get pypy3 dev files needed to install pypy3 packages.
RUN apt-get install -y pypy3-dev

# Clone the python3 version of ViperMonkey so we can install the
# pypy3 ViperMonkey packages.
RUN cd /opt/ && git clone -b python3_port https://github.com/kirk-sayre-work/ViperMonkey.git

# Install the pypy3 ViperMonkey packages.
RUN pypy3 -m pip install -r /opt/ViperMonkey/requirements.txt

# Get rid of old get-pip.py file.
RUN rm -f /opt/get-pip.py

# For testing. Keeps the container up and running so you can shell
# into it. Comment this out for prod.
#RUN bash -c "while true; do echo 'Keeping container alive ...'; sleep 10; done"
