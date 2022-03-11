FROM python:3.10.2-slim-buster

# Installing Oracle instant client
WORKDIR    /opt/oracle
RUN        apt-get update && apt-get install -y libaio1 wget unzip \
            && wget https://download.oracle.com/otn_software/linux/instantclient/instantclient-basiclite-linuxx64.zip \
            && unzip instantclient-basiclite-linuxx64.zip \
            && rm -f instantclient-basiclite-linuxx64.zip \
            && cd /opt/oracle/instantclient* \
            && rm -f *jdbc* *occi* *mysql* *README *jar uidrvci genezi adrci \
            && echo /opt/oracle/instantclient* > /etc/ld.so.conf.d/oracle-instantclient.conf \
            && ldconfig

WORKDIR    /app
# Copy my project folder content into /app container directory
COPY       . .
RUN         python -m pip install --upgrade pip \
            &&  pip3 install -r requirements.txt \
            && pip3 list

#RUN        pip3 install pipenv
#RUN        pipenv install
EXPOSE     5000
# For this statement to work you need to add the next two lines into Pipfilefile
# [scripts]
# server = "python manage.py runserver 0.0.0.0:8000"
#ENTRYPOINT ["pipenv", "run", "server"]

CMD [ "python3", "-m" , "flask", "run", "--host=0.0.0.0"]