# Step 1: Menggunakan image Nginx Alpine yang sangat kecil
FROM nginx:alpine

# Step 2: Hapus default config nginx
RUN rm /etc/nginx/conf.d/default.conf

# Step 3: Copy file project SAPUI5 ke folder html nginx
# Pastikan file index.html ada di root folder apps/sapui5
COPY . /usr/share/nginx/html

# Step 4: Tambahkan konfigurasi sederhana agar routing UI5 lancar
COPY <<EOF /etc/nginx/conf.d/default.conf
server {
    listen 80;
    server_name localhost;
    location / {
        root   /usr/share/nginx/html;
        index  index.html index.htm;
        try_files \$uri \$uri/ /index.html;
    }
}
EOF

EXPOSE 80
CMD ["nginx", "-g", "daemon off;"]