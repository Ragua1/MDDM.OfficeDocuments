# Temporal Compute Orchestrator - Technical Specification

## 1. Project Overview

### 1.1 Project Identity
- **Name**: temporal-compute-orchestrator
- **Primary Language**: .NET (C#)
- **Secondary Language**: Python
- **Orchestration Engine**: Temporal
- **License**: Apache-2.0 or MIT

### 1.2 Goal
Self-hosted orchestrator for scheduling and executing compute/generation jobs on heterogeneous worker nodes using Temporal.

## 2. Context and Problem Statement

### 2.1 Problem
User owns multiple machines (office PC, home PC, Raspberry Pi) with idle compute capacity. They want to remotely schedule, manage and monitor compute/generation jobs across these machines. Nodes can be offline, busy or overloaded. Jobs must be queued, retried and routed dynamically.

### 2.2 Key Constraints
- **Self-hosted**: Complete control over infrastructure and data
- **Open source friendly**: Permissive licensing and community-driven development
- **Workers connect outbound**: NAT-friendly architecture - no inbound ports required
- **Durable execution**: Job state persists across failures and restarts

## 3. Business Requirements

### 3.1 Job Management

#### 3.1.1 Capabilities
The system must support the following job management operations:
- **Create job**: Submit new jobs to the system
- **Cancel job**: Terminate running or queued jobs
- **Query job status**: Get current state and progress of a job
- **List jobs**: Enumerate all jobs with filtering capabilities
- **Retrieve job result metadata**: Access execution results and artifacts

#### 3.1.2 Job Attributes
Each job contains the following attributes:

| Attribute | Type | Description |
|-----------|------|-------------|
| `jobType` | enum | Type of job: `compute`, `generate`, or `script` |
| `priority` | enum | Execution priority: `low`, `normal`, or `high` |
| `capabilities` | object | Required worker capabilities (cpu, gpu, ram, arch, os) |
| `maxRuntime` | duration | Maximum allowed execution time |
| `retryPolicy` | object | Configuration for automatic retries |
| `timeout` | duration | Overall job timeout including retries |
| `payload` | object | Job-specific data and parameters |

### 3.2 Scheduling

#### 3.2.1 Requirements
- Jobs are queued when no suitable worker is available
- Workers actively pull tasks from assigned queues
- System tolerates worker disconnects and reconnects
- Concurrency limits enforced per worker
- Priority-based scheduling ensures high-priority jobs execute first

### 3.3 Workers

#### 3.3.1 Worker Types
The system supports heterogeneous worker nodes:

| Type | Architecture | Characteristics |
|------|-------------|-----------------|
| Office PC | x64 | High computational power, may be intermittently available |
| Home PC | x64 | Moderate power, potentially more stable availability |
| Raspberry Pi | arm64 | Low power, lightweight tasks, always-on potential |

#### 3.3.2 Worker Properties

| Property | Description |
|----------|-------------|
| `workerId` | Unique identifier for the worker |
| `capabilities` | List of supported job types and hardware features |
| `tags` | Location or logical grouping (e.g., home, office) |
| `maxConcurrency` | Maximum number of concurrent jobs |
| `lastHeartbeat` | Timestamp of last activity for health monitoring |

## 4. Non-Functional Requirements

### 4.1 Reliability
- **No job loss on crash or restart**: Job state persisted to durable storage
- **Automatic retries with backoff**: Exponential backoff for transient failures
- **Idempotent job submission**: Duplicate submissions don't create multiple jobs

### 4.2 Observability
- **Structured logging**: Machine-parseable logs for analysis
- **Job and worker visibility**: Real-time status via API and UI
- **Temporal Web UI support**: Integration with Temporal's built-in UI

### 4.3 Security
- **API protected by token**: Bearer token authentication for REST API
- **TLS-ready Temporal configuration**: Encrypted communication support

## 5. Architecture

### 5.1 Components

#### 5.1.1 Temporal Server
**Purpose**: Durable workflow orchestration engine

**Deployment Options**:
- Development: Temporal CLI dev server (`temporal server start-dev`)
- Production: Docker Compose with PostgreSQL backend

**References**:
- [Temporal CLI Documentation](https://docs.temporal.io/cli)
- [Temporal Server Samples](https://github.com/temporalio/samples-server)

#### 5.1.2 Orchestrator API
**Technology**: ASP.NET Core (.NET 8)

**Responsibilities**:
- Expose REST API for job management
- Validate job requests against schema
- Start Temporal workflows for approved jobs
- Query and expose job and worker state

**Endpoints**:

| Method | Path | Description |
|--------|------|-------------|
| POST | `/jobs` | Create a new job |
| GET | `/jobs/{jobId}` | Get job status and details |
| POST | `/jobs/{jobId}/cancel` | Cancel a running or queued job |
| GET | `/workers` | List registered workers and their status |

#### 5.1.3 Temporal Workflows

**Workflow**: `JobWorkflow`
- Orchestrates the complete lifecycle of a job
- Handles routing to appropriate task queue
- Manages retries and timeouts
- Tracks job state transitions

**Activities**:
- `DispatchActivity`: Route job to appropriate task queue based on requirements
- `ExecuteActivity`: Execute the actual job on a worker
- `CollectResultActivity`: Gather and store execution results

**Requirements**:
- Deterministic logic (no random, time, or I/O in workflow code)
- Configurable retry and timeout policies
- Graceful cancellation support

#### 5.1.4 Workers

**5.1.4.1 .NET Worker**
- **Technology**: .NET Worker Service
- **Responsibilities**:
  - Poll Temporal task queues for work
  - Execute jobs according to type and payload
  - Respect configured concurrency limits
  - Report execution results and errors back to Temporal

**5.1.4.2 Python Worker (Optional)**
- **Technology**: Python with Temporal Python SDK
- **Responsibilities**:
  - Execute Python-based compute jobs
  - Serve as lightweight agent on ARM devices (Raspberry Pi)
  - Interface with Python-specific libraries and tools

### 5.2 Task Queue Strategy

#### 5.2.1 Queue Definitions
The system uses multiple task queues for job routing:

| Queue Name | Purpose |
|------------|---------|
| `cpu-x64` | General-purpose compute on x64 architecture |
| `gpu-x64` | GPU-accelerated tasks on x64 architecture |
| `arm64-light` | Lightweight tasks on ARM64 devices |
| `office` | Jobs specifically for office location workers |
| `home` | Jobs specifically for home location workers |

#### 5.2.2 Routing Rules
Job routing follows this decision tree:
1. If `requiresGpu == true` → `gpu-x64`
2. Else if `arch == arm64` → `arm64-light`
3. Else if `location == office` → `office`
4. Else if `location == home` → `home`
5. Else → `cpu-x64` (default)

### 5.3 Retry and Timeout Policy

#### 5.3.1 Activity Retry Policy
```
maxAttempts: 3
backoff: exponential (e.g., 1s, 2s, 4s)
```

#### 5.3.2 Timeout Configuration

| Timeout Type | Purpose |
|-------------|----------|
| `scheduleToStart` | Prevents infinite queue wait; fails if no worker picks up task |
| `startToClose` | Maximum runtime for a single activity execution |

## 6. Repository Structure

### 6.1 Infrastructure
```
temporal/
  docker-compose.yml    # Production-ready Temporal setup with PostgreSQL
```

### 6.2 Source Code
```
src/
  Orchestrator.Api/        # ASP.NET Core REST API
  Orchestrator.Domain/     # Domain models and business logic
  Orchestrator.Temporal/   # Temporal workflows and activities
  Worker.DotNet/          # .NET worker service
  Worker.Python/          # Python worker (optional)
```

### 6.3 Documentation
```
docs/
  architecture.md      # Detailed architecture diagrams and decisions
  task-queues.md      # Task queue configuration and routing logic
  security.md         # Security considerations and best practices
```

## 7. Technology References

### 7.1 Temporal SDKs
- [Temporal .NET SDK](https://github.com/temporalio/sdk-dotnet)
- [Temporal .NET Samples](https://github.com/temporalio/samples-dotnet)
- [Temporal Python Samples](https://github.com/temporalio/samples-python)

### 7.2 Documentation
- [Temporal Documentation](https://docs.temporal.io)

## 8. Definition of Done

The project is considered complete when the following criteria are met:

- [ ] Temporal dev server runs locally without errors
- [ ] API can create jobs and returns job ID
- [ ] API can cancel jobs successfully
- [ ] Jobs execute on workers via task queues
- [ ] Jobs survive worker restart and resume execution
- [ ] README explains setup and usage clearly
- [ ] At least one .NET worker can be started and processes jobs
- [ ] Basic integration test demonstrates end-to-end flow

## 9. Implementation Phases

### Phase 1: Foundation
- Set up Temporal server (dev mode)
- Implement basic API with job creation endpoint
- Create simple JobWorkflow with one activity
- Implement basic .NET worker

### Phase 2: Core Features
- Add job cancellation
- Implement task queue routing
- Add retry and timeout policies
- Implement job status queries

### Phase 3: Production Readiness
- Add authentication/authorization
- Docker Compose for Temporal with PostgreSQL
- Implement Python worker
- Add comprehensive logging and monitoring

### Phase 4: Advanced Features
- Priority-based scheduling
- Worker capability matching
- Advanced retry strategies
- Web dashboard (optional)

## 10. Future Extensions

Potential enhancements beyond initial release:

- **GPU-aware scheduling**: Detect and utilize GPU resources
- **Idle detection on workers**: Only run jobs when machine is idle
- **Cost-aware scheduling**: Consider energy costs or machine availability
- **Web dashboard**: Visual interface for job and worker management
- **Job dependencies**: Support for DAG-based job execution
- **Resource quotas**: Limit resource usage per user or project
- **Job templates**: Predefined job configurations for common tasks
