FROM redmine:3.3

RUN set -x && \
    sed -i \
      -e 's/^\([ \t]*bundle install .*\)$/#\1/' \
      -e '/^[ \t]*set -- gosu redmine .*$/i export RAILS_RELATIVE_URL_ROOT=\/redmine' \
      /docker-entrypoint.sh && \
    sed -i \
      -e '/run RedmineApp::Application/i map Rails.application.config.relative_url_root || "/" do' \
      -e '/run RedmineApp::Application/a end' \
      /usr/src/redmine/config.ru && \
    echo done.
