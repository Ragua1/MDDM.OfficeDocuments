# Temporal Compute Orchestrator

**Self-hosted orchestration platform for scheduling compute and generation jobs across distributed machines using Temporal.**

[![License](https://img.shields.io/badge/license-Apache--2.0%20OR%20MIT-blue.svg)](LICENSE.md)
[![.NET](https://img.shields.io/badge/.NET-8.0-purple.svg)](https://dotnet.microsoft.com/)
[![Temporal](https://img.shields.io/badge/Temporal-latest-00ADD8.svg)](https://temporal.io/)

## Overview

Temporal Compute Orchestrator enables you to harness the idle computing power of your multiple machines - whether it's your office PC, home workstation, or a Raspberry Pi - by providing a robust job scheduling and execution platform. Built on [Temporal](https://temporal.io/), it ensures durable, reliable job execution even when workers go offline or encounter failures.

## ✨ Features

- **🔄 Durable Job Execution**: Jobs survive crashes, restarts, and network failures
- **👷 Worker-Based Scheduling**: Distributed workers pull tasks from queues based on capabilities
- **📡 Offline-Tolerant Nodes**: Workers can disconnect and reconnect without job loss
- **⚡ Priority Queues**: High-priority jobs get executed first
- **🖥️ Heterogeneous Hardware Support**: Run jobs on x64, ARM64, GPU-enabled, or specialized hardware
- **🔍 Built-in Observability**: Monitor jobs and workers via Temporal Web UI and REST API
- **🔐 Secure by Default**: Token-based API authentication and TLS-ready configuration

## 🏗️ Architecture

```
┌─────────────────┐
│   REST API      │  ← Submit jobs, query status
│ (ASP.NET Core)  │
└────────┬────────┘
         │
    ┌────▼────────────────┐
    │  Temporal Server    │  ← Workflow orchestration
    │  (Durable storage)  │
    └────────┬────────────┘
             │
       ┌─────┴──────┬──────────┬──────────┐
       ▼            ▼          ▼          ▼
  ┌─────────┐ ┌─────────┐ ┌─────────┐ ┌─────────┐
  │ Worker  │ │ Worker  │ │ Worker  │ │ Worker  │
  │(Office) │ │ (Home)  │ │  (RPi)  │ │ (Cloud) │
  │  x64    │ │  x64    │ │ arm64   │ │   GPU   │
  └─────────┘ └─────────┘ └─────────┘ └─────────┘
```

## 🚀 Quick Start

### Prerequisites

- [.NET 8.0 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)
- [Temporal CLI](https://docs.temporal.io/cli) (for development)
- (Optional) [Python 3.9+](https://www.python.org/) for Python workers
- (Optional) [Docker](https://www.docker.com/) for production Temporal deployment

### 1. Start Temporal Server

For development, use the Temporal dev server:

```bash
temporal server start-dev
```

For production, use Docker Compose (see `temporal/docker-compose.yml`):

```bash
cd temporal
docker-compose up -d
```

### 2. Run the Orchestrator API

```bash
dotnet run --project src/Orchestrator.Api
```

The API will be available at `http://localhost:5000` (or as configured).

### 3. Run a Worker

Start one or more workers on your machines:

**On your x64 machines:**
```bash
dotnet run --project src/Worker.DotNet
```

**On Raspberry Pi or ARM devices (optional Python worker):**
```bash
cd src/Worker.Python
pip install -r requirements.txt
python worker.py
```

### 4. Submit a Job

Submit your first job via the REST API:

```bash
curl -X POST http://localhost:5000/jobs \
  -H "Content-Type: application/json" \
  -H "Authorization: Bearer YOUR_TOKEN" \
  -d '{
    "jobType": "compute",
    "priority": "normal",
    "capabilities": {
      "cpu": true,
      "arch": "x64"
    },
    "payload": {
      "command": "python script.py"
    }
  }'
```

Response:
```json
{
  "jobId": "job_abc123",
  "status": "queued",
  "createdAt": "2026-01-25T22:00:00Z"
}
```

### 5. Check Job Status

```bash
curl http://localhost:5000/jobs/job_abc123 \
  -H "Authorization: Bearer YOUR_TOKEN"
```

Response:
```json
{
  "jobId": "job_abc123",
  "status": "completed",
  "workerId": "worker-office-pc",
  "startedAt": "2026-01-25T22:00:05Z",
  "completedAt": "2026-01-25T22:05:30Z",
  "result": {
    "exitCode": 0,
    "output": "Task completed successfully"
  }
}
```

## 📚 Usage Examples

### Submit a High-Priority GPU Job

```bash
curl -X POST http://localhost:5000/jobs \
  -H "Content-Type: application/json" \
  -H "Authorization: Bearer YOUR_TOKEN" \
  -d '{
    "jobType": "generate",
    "priority": "high",
    "capabilities": {
      "gpu": true,
      "arch": "x64"
    },
    "maxRuntime": "1h",
    "payload": {
      "model": "stable-diffusion",
      "prompt": "A beautiful sunset over mountains"
    }
  }'
```

### Submit a Script to Raspberry Pi

```bash
curl -X POST http://localhost:5000/jobs \
  -H "Content-Type: application/json" \
  -H "Authorization: Bearer YOUR_TOKEN" \
  -d '{
    "jobType": "script",
    "priority": "low",
    "capabilities": {
      "arch": "arm64"
    },
    "payload": {
      "script": "#!/bin/bash\necho Hello from RPi\n"
    }
  }'
```

### Cancel a Running Job

```bash
curl -X POST http://localhost:5000/jobs/job_abc123/cancel \
  -H "Authorization: Bearer YOUR_TOKEN"
```

### List All Workers

```bash
curl http://localhost:5000/workers \
  -H "Authorization: Bearer YOUR_TOKEN"
```

Response:
```json
{
  "workers": [
    {
      "workerId": "worker-office-pc",
      "status": "active",
      "capabilities": ["cpu", "gpu"],
      "architecture": "x64",
      "tags": ["office"],
      "maxConcurrency": 4,
      "currentJobs": 1,
      "lastHeartbeat": "2026-01-25T22:45:00Z"
    },
    {
      "workerId": "worker-rpi-home",
      "status": "active",
      "capabilities": ["cpu"],
      "architecture": "arm64",
      "tags": ["home"],
      "maxConcurrency": 1,
      "currentJobs": 0,
      "lastHeartbeat": "2026-01-25T22:44:55Z"
    }
  ]
}
```

## 🔧 Configuration

### Orchestrator API Configuration

Edit `src/Orchestrator.Api/appsettings.json`:

```json
{
  "Temporal": {
    "Host": "localhost:7233",
    "Namespace": "default"
  },
  "Authentication": {
    "ApiKey": "your-secret-token-here"
  },
  "Jobs": {
    "DefaultTimeout": "1h",
    "MaxRetries": 3
  }
}
```

### Worker Configuration

Edit `src/Worker.DotNet/appsettings.json`:

```json
{
  "Temporal": {
    "Host": "localhost:7233",
    "Namespace": "default"
  },
  "Worker": {
    "WorkerId": "worker-office-pc",
    "TaskQueues": ["cpu-x64", "office"],
    "MaxConcurrency": 4,
    "Capabilities": {
      "cpu": true,
      "gpu": true,
      "arch": "x64"
    },
    "Tags": ["office"]
  }
}
```

## 📋 Job Types

### Compute Jobs
Execute computational tasks (data processing, analysis, simulations):
```json
{
  "jobType": "compute",
  "payload": {
    "command": "python compute.py --input data.csv"
  }
}
```

### Generate Jobs
Create content (images, videos, models):
```json
{
  "jobType": "generate",
  "payload": {
    "generator": "image",
    "parameters": { "width": 1024, "height": 768 }
  }
}
```

### Script Jobs
Run arbitrary scripts or commands:
```json
{
  "jobType": "script",
  "payload": {
    "script": "#!/bin/bash\nbackup.sh"
  }
}
```

## 🎯 Task Queue Routing

Jobs are automatically routed to appropriate task queues based on their requirements:

| Job Requirement | Target Queue |
|----------------|--------------|
| `requiresGpu: true` | `gpu-x64` |
| `arch: arm64` | `arm64-light` |
| `location: office` | `office` |
| `location: home` | `home` |
| Default | `cpu-x64` |

Workers subscribe to specific queues based on their capabilities.

## 🔄 Retry and Timeout Policies

### Default Retry Policy
- **Max Attempts**: 3
- **Backoff**: Exponential (1s, 2s, 4s)
- **Non-Retryable Errors**: Validation errors, missing dependencies

### Default Timeouts
- **Schedule to Start**: 5 minutes (fail if no worker available)
- **Start to Close**: 1 hour (configurable per job)

## 🛠️ Development

### Project Structure

```
temporal-compute-orchestrator/
├── temporal/
│   └── docker-compose.yml          # Production Temporal setup
├── src/
│   ├── Orchestrator.Api/           # REST API
│   ├── Orchestrator.Domain/        # Domain models
│   ├── Orchestrator.Temporal/      # Workflows & Activities
│   ├── Worker.DotNet/             # .NET worker service
│   └── Worker.Python/             # Python worker (optional)
├── docs/
│   ├── architecture.md            # Architecture details
│   ├── task-queues.md            # Queue configuration
│   └── security.md               # Security best practices
├── spec.md                        # Technical specification
└── README.md                      # This file
```

### Building the Project

```bash
# Restore dependencies
dotnet restore

# Build all projects
dotnet build

# Run tests
dotnet test
```

### Running Locally

See the [Quick Start](#-quick-start) section above.

### Accessing Temporal Web UI

Once the Temporal server is running, access the Web UI at:
- Dev server: `http://localhost:8233`
- Docker: `http://localhost:8080` (as configured in docker-compose.yml)

The UI provides visibility into:
- Workflow executions
- Activity history
- Job status and timeline
- Worker task queues

## 🔐 Security

### API Authentication

All API endpoints require a Bearer token:

```bash
curl -H "Authorization: Bearer YOUR_SECRET_TOKEN" \
  http://localhost:5000/jobs
```

Configure the token in `appsettings.json` or via environment variable:

```bash
export ORCHESTRATOR_API_KEY="your-secret-token"
```

### TLS Configuration

For production deployments, enable TLS for:
1. **API endpoints**: Configure HTTPS in ASP.NET Core
2. **Temporal connections**: Use mTLS for worker-to-server communication

See `docs/security.md` for detailed security guidance.

## 🚀 Deployment

### Production Checklist

- [ ] Deploy Temporal with PostgreSQL (see `temporal/docker-compose.yml`)
- [ ] Enable TLS for all connections
- [ ] Configure proper authentication tokens
- [ ] Set up monitoring and alerting
- [ ] Configure backup for Temporal database
- [ ] Deploy API behind reverse proxy (nginx, Traefik)
- [ ] Configure firewall rules (only workers need outbound access)
- [ ] Set resource limits for workers

### Docker Deployment

```bash
# Start Temporal and PostgreSQL
cd temporal
docker-compose up -d

# Build and run API
docker build -t orchestrator-api -f src/Orchestrator.Api/Dockerfile .
docker run -p 5000:5000 orchestrator-api

# Build and run worker
docker build -t orchestrator-worker -f src/Worker.DotNet/Dockerfile .
docker run orchestrator-worker
```

## 🔮 Future Extensions

Planned enhancements for future releases:

- **GPU-aware scheduling**: Automatically detect and utilize GPU resources
- **Idle detection**: Only run jobs when worker machines are idle
- **Cost-aware scheduling**: Consider energy costs and optimal execution times
- **Web dashboard**: Rich UI for job and worker management
- **Job dependencies**: Support for DAG-based workflows
- **Resource quotas**: Per-user or per-project limits
- **Job templates**: Reusable configurations for common tasks
- **Metrics and analytics**: Historical performance tracking

## 📖 Documentation

- [Technical Specification](spec.md) - Detailed system design and requirements
- [Architecture Guide](docs/architecture.md) - Component details and diagrams
- [Task Queue Configuration](docs/task-queues.md) - Queue routing and setup
- [Security Best Practices](docs/security.md) - Hardening and authentication
- [Temporal Documentation](https://docs.temporal.io) - Official Temporal docs

## 🤝 Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📄 License

This project is dual-licensed under:
- [Apache License 2.0](LICENSE-APACHE)
- [MIT License](LICENSE-MIT)

You may choose either license for your use.

## 🙏 Acknowledgments

- [Temporal](https://temporal.io/) - Durable execution engine
- [.NET Foundation](https://dotnetfoundation.org/) - .NET ecosystem
- All contributors who help improve this project

## 📞 Support

- **Issues**: [GitHub Issues](https://github.com/your-org/temporal-compute-orchestrator/issues)
- **Discussions**: [GitHub Discussions](https://github.com/your-org/temporal-compute-orchestrator/discussions)
- **Documentation**: [Project Wiki](https://github.com/your-org/temporal-compute-orchestrator/wiki)

---

**Note**: This is a self-hosted solution. You maintain complete control over your infrastructure, data, and job execution.
