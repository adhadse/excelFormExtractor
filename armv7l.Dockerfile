FROM ghcr.io/bjia56/armv7l-wheel-builder:main@sha256:7664c40814ba3d9d546d61b4cab16608b86d978f21204a1e59e2efdf1437e8fd

ARG RELEASE_VERSION
ENV RELEASE_VERSION=${RELEASE_VERSION}

RUN go install github.com/go-python/gopy@v0.4.10 && \
    go install golang.org/x/tools/cmd/goimports@v0.16.0

RUN mkdir build
WORKDIR build
COPY . .

RUN for ver in {3.10,3.11,3.12};  \
    do  \
    mkdir -p build${ver} && \
    python${ver} setup.py bdist_wheel --dist-dir build${ver} && \
    auditwheel repair build${ver}/*armv7l.whl --wheel-dir build${ver}/wheelhouse && \
    ls build${ver} ;\
    done

RUN mkdir -p /export && \
    cp build*/wheelhouse/*.whl /export
