"""
Batch processor for approval instances.

Handles batch downloading of attachments from multiple approval instances.
"""

from typing import List, Callable, Optional


class BatchProcessor:
    """
    Process multiple approval instances in batches.
    """

    def __init__(self, max_workers: int = 3):
        """
        Initialize batch processor.

        Args:
            max_workers: Maximum concurrent download workers.
        """
        self.max_workers = max_workers

    def process_instances(
        self,
        instance_codes: List[str],
        progress_callback: Optional[Callable[[int, int], None]] = None,
    ) -> List[str]:
        """
        Process a list of instance codes, downloading attachments.

        Args:
            instance_codes: List of instance codes to process.
            progress_callback: Optional callback(completed, total).

        Returns:
            List of downloaded file paths.
        """
        # TODO: Implement
        pass

    def process_single(self, instance_code: str) -> List[str]:
        """
        Process a single instance and download its attachments.

        Args:
            instance_code: Instance code to process.

        Returns:
            List of downloaded file paths.
        """
        # TODO: Implement
        pass
