FROM nginx:alpine

# 1. Hapus semua file default nginx agar tidak bentrok
RUN rm -rf /usr/share/nginx/html/*

# 2. Copy isi folder public ke root web server nginx
# Ini akan memindahkan index.html dari public/ ke /usr/share/nginx/html/
COPY public/ /usr/share/nginx/html/

# 3. Beri izin akses folder
RUN chmod -R 755 /usr/share/nginx/html

EXPOSE 80
CMD ["nginx", "-g", "daemon off;"]