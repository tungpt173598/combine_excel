FROM php:8.1-fpm

# Các bước cài đặt các gói phụ thuộc và công cụ cần thiết
RUN apt-get update && apt-get install -y \
    libzip-dev \
    && docker-php-ext-install zip \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

# Cấu hình tùy chỉnh của PHP
COPY php.ini /usr/local/etc/php/conf.d/custom.ini

# Thư mục làm việc mặc định của ứng dụng
WORKDIR /var/www/html

# Mở cổng mặc định cho PHP-FPM (nếu cần thiết)
EXPOSE 9000

# Lệnh khởi chạy của Docker container
CMD ["php-fpm"]
