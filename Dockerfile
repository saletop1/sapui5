FROM nginx:alpine

# 1. Hapus total isi folder dan config default
RUN rm -rf /usr/share/nginx/html/*
RUN rm /etc/nginx/conf.d/default.conf

# 2. Copy file project (folder public)
COPY public/ /usr/share/nginx/html/

# 3. Copy config nginx yang baru kita buat tadi
COPY nginx.conf /etc/nginx/conf.d/default.conf

# 4. Beri izin
RUN chmod -R 755 /usr/share/nginx/html

EXPOSE 80
CMD ["nginx", "-g", "daemon off;"]