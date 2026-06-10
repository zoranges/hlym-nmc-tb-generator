# HLYM NMC TB Auto Generator — 部署文档

## 镜像信息

| 项目 | 值 |
|---|---|
| 镜像 | `zosic/tb-generator:latest` |
| 端口 | `8080` |
| 数据挂载 | `/data/data.xlsx`（可选） |
| 输出挂载 | `/app/output`（可选） |

## 快速部署

```bash
# 1. 拉取镜像
docker pull zosic/tb-generator:latest

# 2. 启动容器（使用内置数据）
docker run -d \
  --name tb-generator \
  --restart unless-stopped \
  -p 8080:8080 \
  zosic/tb-generator:latest

# 3. 验证
curl http://localhost:8080/api/health
```

打开浏览器访问 `http://<服务器IP>:8080`

## 挂载自定义数据文件

如果你有自己的 Excel 数据源文件：

```bash
docker run -d \
  --name tb-generator \
  --restart unless-stopped \
  -p 8080:8080 \
  -v /path/to/your/data.xlsx:/data/data.xlsx \
  -v /path/to/output:/app/output \
  zosic/tb-generator:latest
```

## 反向代理（Nginx）

```nginx
server {
    listen 80;
    server_name your-domain.com;

    client_max_body_size 50m;

    location / {
        proxy_pass http://127.0.0.1:8080;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
    }
}
```

## 容器管理

```bash
# 查看日志
docker logs -f tb-generator

# 重启
docker restart tb-generator

# 停止
docker stop tb-generator

# 删除
docker rm -f tb-generator

# 更新镜像
docker pull zosic/tb-generator:latest
docker rm -f tb-generator
# 然后重新执行上面的 docker run 命令
```

## 安全建议

1. 生产环境建议加 HTTPS（Cloudflare / Let's Encrypt）
2. 如需鉴权，可在 Nginx 层加 Basic Auth
3. 50MB 上传限制可通过 Nginx `client_max_body_size` 调整
