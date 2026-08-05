# P. Vertex BHT dashboard. index.html is fully self-contained -- the 500 survey
# records, Chart.js and the brand logos are all inlined into the file -- so this
# is a pure static serve with no build step and no runtime data dependency.
# dashboard.html is currently an identical copy, kept so existing links to it
# keep resolving.
FROM nginx:1.30.4-alpine

COPY docker/conf.d/default.conf /etc/nginx/conf.d/default.conf
COPY index.html dashboard.html /var/www/public/

EXPOSE 80
CMD ["nginx", "-g", "daemon off;"]
