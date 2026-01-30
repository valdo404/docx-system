FROM --platform=$BUILDPLATFORM mcr.microsoft.com/dotnet/sdk:10.0-preview AS build
ARG TARGETARCH

WORKDIR /src

# Copy project files first for layer caching
COPY src/DocxMcp/DocxMcp.csproj src/DocxMcp/
RUN dotnet restore src/DocxMcp/DocxMcp.csproj \
    -a $TARGETARCH

# Copy everything and publish
COPY . .
RUN dotnet publish src/DocxMcp/DocxMcp.csproj \
    --configuration Release \
    --no-restore \
    -a $TARGETARCH \
    -o /app

# Runtime: minimal image with only the binary
# The runtime-deps image already provides an 'app' user/group
FROM mcr.microsoft.com/dotnet/runtime-deps:10.0-preview AS runtime

WORKDIR /app
COPY --from=build /app .

USER app

ENTRYPOINT ["./docx-mcp"]
