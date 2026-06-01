"""워크플로우 등록부. 새 템플릿 = 모듈 1개 추가 + @register 데코레이터."""
from typing import Dict, Type
from core.base import Workflow

_registry: Dict[str, Type[Workflow]] = {}


def register(cls: Type[Workflow]):
    """워크플로우 등록 데코레이터. cls.name 을 키로 사용."""
    _registry[cls.name] = cls
    return cls


def list_workflows() -> list[str]:
    return list(_registry.keys())


def get_workflow(name: str) -> Workflow:
    if name not in _registry:
        raise KeyError(f"워크플로우 없음: '{name}'. 등록된 목록: {list_workflows()}")
    return _registry[name]()
