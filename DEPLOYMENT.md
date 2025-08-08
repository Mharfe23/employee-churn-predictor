# Deployment Guide

This guide covers deploying the Employee Churn Predictor Excel Add-in to different platforms and environments.

## 🚀 Deployment Options

### 1. Local Development (Current Setup)
- **Use Case**: Development and testing
- **Setup**: Use `start_services.ps1` or manual setup
- **Access**: Localhost only

### 2. Production Deployment
- **Use Case**: Production use in organizations
- **Setup**: Deploy to cloud services
- **Access**: Internet accessible

### 3. Enterprise Deployment
- **Use Case**: Large organizations with security requirements
- **Setup**: On-premises deployment
- **Access**: Internal network only

## ☁️ Cloud Deployment

### Option A: Azure Deployment

#### Prerequisites
- Azure subscription
- Azure CLI installed

#### Steps

1. **Deploy Flask Server to Azure App Service:**
   ```bash
   # Login to Azure
   az login
   
   # Create resource group
   az group create --name churn-predictor-rg --location eastus
   
   # Create App Service plan
   az appservice plan create --name churn-predictor-plan --resource-group churn-predictor-rg --sku B1
   
   # Create web app
   az webapp create --name churn-predictor-api --resource-group churn-predictor-rg --plan churn-predictor-plan --runtime "PYTHON|3.9"
   
   # Deploy Flask app
   cd flask_server
   az webapp deployment source config-local-git --name churn-predictor-api --resource-group churn-predictor-rg
   git remote add azure <azure-git-url>
   git push azure main
   ```

2. **Deploy Excel Add-in to Azure Static Web Apps:**
   ```bash
   # Build the add-in
   cd employe_ml_excel_addin
   npm run build
   
   # Deploy to Azure Static Web Apps
   az staticwebapp create --name churn-predictor-addin --resource-group churn-predictor-rg --source .
   ```

3. **Update manifest with production URLs:**
   ```xml
   <SourceLocation DefaultValue="https://churn-predictor-addin.azurestaticapps.net/taskpane.html"/>
   ```

### Option B: AWS Deployment

#### Prerequisites
- AWS account
- AWS CLI configured

#### Steps

1. **Deploy Flask Server to AWS Lambda:**
   ```bash
   # Create deployment package
   cd flask_server
   pip install -r requirements.txt -t .
   zip -r lambda-deployment.zip .
   
   # Deploy to Lambda
   aws lambda create-function \
     --function-name churn-predictor-api \
     --runtime python3.9 \
     --handler app.lambda_handler \
     --zip-file fileb://lambda-deployment.zip
   ```

2. **Deploy Excel Add-in to S3:**
   ```bash
   # Build the add-in
   cd employe_ml_excel_addin
   npm run build
   
   # Deploy to S3
   aws s3 sync dist/ s3://churn-predictor-addin --delete
   ```

### Option C: Heroku Deployment

#### Steps

1. **Deploy Flask Server:**
   ```bash
   cd flask_server
   heroku create churn-predictor-api
   git init
   git add .
   git commit -m "Initial deployment"
   git push heroku main
   ```

2. **Deploy Excel Add-in:**
   ```bash
   cd employe_ml_excel_addin
   npm run build
   # Deploy dist/ folder to any static hosting service
   ```

## 🏢 Enterprise Deployment

### On-Premises Setup

1. **Deploy Flask Server:**
   ```bash
   # On your server
   cd flask_server
   python -m venv venv
   source venv/bin/activate
   pip install -r requirements.txt
   
   # Use gunicorn for production
   pip install gunicorn
   gunicorn -w 4 -b 0.0.0.0:5000 app:app
   ```

2. **Deploy Excel Add-in:**
   ```bash
   # Build for production
   cd employe_ml_excel_addin
   npm run build
   
   # Serve with nginx or Apache
   # Copy dist/ contents to web server directory
   ```

3. **Configure Reverse Proxy (nginx example):**
   ```nginx
   server {
       listen 80;
       server_name your-domain.com;
       
       location / {
           root /var/www/churn-predictor-addin;
           index index.html;
       }
       
       location /api/ {
           proxy_pass http://localhost:5000/;
           proxy_set_header Host $host;
           proxy_set_header X-Real-IP $remote_addr;
       }
   }
   ```

## 📦 Distribution Methods

### 1. Shared Network Drive
- Place manifest file on shared drive
- Users can sideload from network location
- Update manifest URLs to point to internal server

### 2. SharePoint/Teams
- Upload manifest to SharePoint
- Users can add from SharePoint location
- Integrate with Microsoft 365

### 3. Intune/SCCM
- Package add-in for enterprise deployment
- Deploy through Microsoft Intune
- Manage updates centrally

## 🔧 Configuration Management

### Environment Variables

Create `.env` files for different environments:

```bash
# Development
FLASK_ENV=development
FLASK_DEBUG=True
CORS_ORIGINS=http://localhost:3001

# Production
FLASK_ENV=production
FLASK_DEBUG=False
CORS_ORIGINS=https://your-domain.com
```

### Update Scripts

Create deployment scripts for each environment:

```powershell
# deploy-dev.ps1
$env:FLASK_ENV="development"
.\start_services.ps1

# deploy-prod.ps1
$env:FLASK_ENV="production"
# Production deployment commands
```

## 🔒 Security Considerations

### Production Security

1. **HTTPS Only**: Always use HTTPS in production
2. **CORS Configuration**: Restrict CORS origins
3. **API Authentication**: Add authentication to Flask API
4. **Rate Limiting**: Implement rate limiting
5. **Input Validation**: Validate all inputs

### Example Security Enhancements

```python
# Add to Flask app.py
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

limiter = Limiter(
    app,
    key_func=get_remote_address,
    default_limits=["200 per day", "50 per hour"]
)

@app.route("/predict", methods=["POST"])
@limiter.limit("10 per minute")
def predict():
    # Existing code
```

## 📊 Monitoring and Logging

### Application Monitoring

1. **Flask Monitoring**: Use Flask-MonitoringDashboard
2. **Error Tracking**: Integrate Sentry
3. **Performance**: Use Application Insights (Azure) or similar

### Logging Configuration

```python
import logging
from logging.handlers import RotatingFileHandler

if not app.debug:
    file_handler = RotatingFileHandler('logs/churn-predictor.log', maxBytes=10240, backupCount=10)
    file_handler.setFormatter(logging.Formatter(
        '%(asctime)s %(levelname)s: %(message)s [in %(pathname)s:%(lineno)d]'
    ))
    file_handler.setLevel(logging.INFO)
    app.logger.addHandler(file_handler)
    app.logger.setLevel(logging.INFO)
    app.logger.info('Churn Predictor startup')
```

## 🚀 Deployment Checklist

- [ ] Environment variables configured
- [ ] HTTPS certificates installed
- [ ] CORS origins updated
- [ ] Database connections configured (if applicable)
- [ ] Logging configured
- [ ] Monitoring set up
- [ ] Backup strategy in place
- [ ] Documentation updated
- [ ] Testing completed
- [ ] Rollback plan prepared

## 🔄 Continuous Deployment

### GitHub Actions for Production

```yaml
# .github/workflows/deploy.yml
name: Deploy to Production

on:
  push:
    tags:
      - 'v*'

jobs:
  deploy:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v3
      - name: Deploy to Azure
        run: |
          # Deployment commands
```

## 📞 Support and Maintenance

### Regular Maintenance Tasks

1. **Security Updates**: Regular dependency updates
2. **Performance Monitoring**: Monitor response times
3. **Backup Verification**: Test backup restoration
4. **User Training**: Provide user documentation
5. **Version Management**: Plan for updates

### Support Documentation

- Create user guides for each deployment method
- Document troubleshooting procedures
- Maintain contact information for support
- Set up monitoring alerts

---

**Need help with deployment?** Check the troubleshooting section or contact your system administrator.
