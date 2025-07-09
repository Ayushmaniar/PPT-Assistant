"""
Phoenix Tracing Configuration for PPT Assistant

This module sets up Phoenix tracing to monitor and debug LLM interactions.
Make sure to add these environment variables to your .env file:

Required Environment Variables:
- PHOENIX_API_KEY: Your Phoenix API key
- PHOENIX_COLLECTOR_ENDPOINT: Phoenix server endpoint (default: http://localhost:6006)
- PHOENIX_PROJECT_NAME: Project name in Phoenix (default: PPT-Assistant)

Optional for Phoenix Cloud:
- For Phoenix Cloud, set PHOENIX_COLLECTOR_ENDPOINT=https://app.phoenix.arize.com
"""

import os
import logging
from typing import Optional
from phoenix.otel import register
from openinference.instrumentation.openai import OpenAIInstrumentor
from opentelemetry import trace
from opentelemetry.exporter.otlp.proto.http.trace_exporter import OTLPSpanExporter
from opentelemetry.sdk.trace import TracerProvider
from opentelemetry.sdk.trace.export import BatchSpanProcessor
from opentelemetry.sdk.resources import Resource
from openinference.semconv.resource import ResourceAttributes

logger = logging.getLogger(__name__)

class PhoenixTracer:
    """Phoenix tracing configuration and management."""
    
    def __init__(self):
        self.tracer_provider = None
        self.tracer = None
        self.is_initialized = False
        
    def initialize_tracing(self) -> bool:
        """
        Initialize Phoenix tracing with environment configuration.
        
        Returns:
            bool: True if initialization was successful, False otherwise
        """
        try:
            # Get Phoenix configuration from environment
            phoenix_api_key = os.getenv("PHOENIX_API_KEY")
            phoenix_endpoint = os.getenv("PHOENIX_COLLECTOR_ENDPOINT", "https://app.phoenix.arize.com")
            project_name = os.getenv("PHOENIX_PROJECT_NAME", "PPT-Assistant")
            
            if not phoenix_api_key:
                logger.warning("PHOENIX_API_KEY not found in environment. Tracing will be disabled.")
                return False
            
            # Configure Phoenix using the register function
            # Use appropriate headers based on endpoint
            if "app.phoenix.arize.com" in phoenix_endpoint:
                # Phoenix Cloud uses api_key header
                headers = {"api_key": phoenix_api_key}
            else:
                # Self-hosted Phoenix uses Authorization header
                headers = {"Authorization": f"Bearer {phoenix_api_key}"}
            
            # Ensure endpoint has the correct path
            if not phoenix_endpoint.endswith("/v1/traces"):
                endpoint = f"{phoenix_endpoint}/v1/traces"
            else:
                endpoint = phoenix_endpoint
            
            self.tracer_provider = register(
                project_name=project_name,
                endpoint=endpoint,
                headers=headers
            )
            
            # Set up OpenAI instrumentation
            OpenAIInstrumentor().instrument()
            
            # Get tracer for manual instrumentation
            self.tracer = trace.get_tracer(__name__)
            
            self.is_initialized = True
            logger.info(f"Phoenix tracing initialized successfully for project: {project_name}")
            logger.info(f"Traces will be sent to: {phoenix_endpoint}")
            
            return True
            
        except Exception as e:
            logger.error(f"Failed to initialize Phoenix tracing: {e}")
            return False
    
    def create_span(self, name: str, **attributes):
        """
        Create a new span for manual tracing.
        
        Args:
            name: Name of the span
            **attributes: Additional attributes to set on the span
            
        Returns:
            Span context manager or None if tracing is not initialized
        """
        if not self.is_initialized or not self.tracer:
            from contextlib import nullcontext
            return nullcontext()
            
        # Start span as current span and return the context manager
        span_context = self.tracer.start_as_current_span(name)
        
        # Set attributes on the span when it's entered
        def _enter_span():
            span = span_context.__enter__()
            for key, value in attributes.items():
                span.set_attribute(key, str(value))
            return span
            
        # Patch the context manager to set attributes
        original_enter = span_context.__enter__
        span_context.__enter__ = _enter_span
            
        return span_context
    
    def add_event(self, name: str, attributes: Optional[dict] = None):
        """
        Add an event to the current span.
        
        Args:
            name: Event name
            attributes: Optional event attributes
        """
        if not self.is_initialized:
            return
            
        current_span = trace.get_current_span()
        if current_span:
            current_span.add_event(name, attributes or {})

# Global tracer instance
phoenix_tracer = PhoenixTracer()

def initialize_phoenix():
    """Initialize Phoenix tracing - call this at application startup."""
    return phoenix_tracer.initialize_tracing()

def trace_function(func_name: str):
    """Decorator to trace function calls."""
    def decorator(func):
        def wrapper(*args, **kwargs):
            if not phoenix_tracer.is_initialized:
                return func(*args, **kwargs)
                
            with phoenix_tracer.create_span(
                f"ppt_assistant.{func_name}",
                function_name=func_name,
                args_count=len(args),
                kwargs_count=len(kwargs)
            ):
                try:
                    result = func(*args, **kwargs)
                    phoenix_tracer.add_event("function_completed", {"success": True})
                    return result
                except Exception as e:
                    phoenix_tracer.add_event("function_error", {"error": str(e), "error_type": type(e).__name__})
                    raise
                    
        return wrapper
    return decorator

def trace_tool_call(tool_name: str, **params):
    """Context manager for tracing tool calls."""
    if not phoenix_tracer.is_initialized:
        from contextlib import nullcontext
        return nullcontext()
        
    return phoenix_tracer.create_span(
        f"tool.{tool_name}",
        tool_name=tool_name,
        **{f"param_{k}": str(v) for k, v in params.items()}
    )

def add_trace_event(event_name: str, **attributes):
    """Add an event to the current trace."""
    phoenix_tracer.add_event(event_name, attributes if attributes else None) 