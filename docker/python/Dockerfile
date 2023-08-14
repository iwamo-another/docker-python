FROM python:3.10
# ARG AWS_ACCESS_KEY_ID
# ARG AWS_SECRET_ACCESS_KEY

# update apt-get
RUN apt-get update -y && apt-get upgrade -y
RUN apt-get install -y vim

# # Install Nodejs 8
# RUN curl -sL https://deb.nodesource.com/setup_18.x | bash -
# RUN apt-get install -y nodejs
# RUN apt-get install -y npm

# Install Python libraries
ADD requirements.txt ./
RUN pip install --upgrade pip && \
    pip install -r requirements.txt

# # install serverless framework
# RUN npm install -g serverless
# RUN npm install --save-dev serverless-step-functions
# RUN npm install --save-dev serverless-pseudo-parameters

# set aws key 
# RUN sls config credentials --provider aws --key $AWS_ACCESS_KEY_ID --secret $AWS_SECRET_ACCESS_KEY

# __pycache__ディレクトリを作成しない
ENV PYTHONDONTWRITEBYTECODE=1

RUN mkdir -p /src

