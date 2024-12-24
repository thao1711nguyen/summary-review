FROM ruby:3.3.5

ENV DEBIAN_FRONTEND noninteractive


RUN apt-get update 
RUN apt-get upgrade -y 
RUN apt-get -y install postgresql postgresql-contrib libpq-dev

WORKDIR /usr/src/app
COPY Gemfile* ./
RUN gem install bundler -v=2.5.16
RUN bundle install

COPY . .
# RUN rails assets:precompile

EXPOSE 80
CMD ["/bin/sh", "-c", "rm -f /usr/src/app/tmp/pids/server.pid && bundle exec rails server -p 80 -b '0.0.0.0'"]
