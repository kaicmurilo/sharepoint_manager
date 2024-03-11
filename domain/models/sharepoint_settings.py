from dataclasses import dataclass


@dataclass
class SharepointSettings:
    site_url: str
    site_path: str
    tenant_id: str
    client_id: str
    client_secret: str
