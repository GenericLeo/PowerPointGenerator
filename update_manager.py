"""
Auto-Update Manager for PowerPoint Generator
Checks GitHub releases for updates and provides download/update functionality
"""

import json
import urllib.request
import urllib.error
import ssl
import webbrowser
import os
import sys
import subprocess
import re
from packaging import version as version_parser
from version import __version__, __github_repo__

try:
    import certifi
except ImportError:
    certifi = None


class UpdateManager:
    """Manages application updates from GitHub releases"""
    
    def __init__(self, github_repo=None):
        """
        Initialize the update manager
        
        Args:
            github_repo: GitHub repository in format 'username/repo'
        """
        self.github_repo = github_repo or __github_repo__
        self.current_version = __version__
        self.api_url = f"https://api.github.com/repos/{self.github_repo}/releases/latest"
        
    def check_for_updates(self, timeout=10):
        """
        Check if a new version is available on GitHub
        
        Args:
            timeout: Request timeout in seconds
            
        Returns:
            dict with keys:
                - update_available (bool): True if update is available
                - latest_version (str): Latest version number
                - download_url (str): URL to download the latest release
                - release_notes (str): Release notes/changelog
                - error (str): Error message if check failed
        """
        try:
            # Make request to GitHub API
            req = urllib.request.Request(
                self.api_url,
                headers={'Accept': 'application/vnd.github.v3+json'}
            )
            
            ssl_context = self._get_ssl_context()

            with urllib.request.urlopen(req, timeout=timeout, context=ssl_context) as response:
                data = json.loads(response.read().decode())
            
            release_notes = data.get('body', 'No release notes available.')
            
            # Find the macOS .dmg or .zip asset
            download_url = None
            assets = data.get('assets', [])

            # Resolve the release version from tag/name/assets in a robust way
            latest_version = self._resolve_release_version(data, assets)
            
            for asset in assets:
                asset_name = asset.get('name', '').lower()
                if asset_name.endswith('.dmg') or asset_name.endswith('.zip'):
                    if 'macos' in asset_name or 'mac' in asset_name or asset_name.endswith('.dmg'):
                        download_url = asset.get('browser_download_url')
                        break
            
            # If no Mac-specific asset found, use first .dmg or .zip
            if not download_url:
                for asset in assets:
                    asset_name = asset.get('name', '').lower()
                    if asset_name.endswith('.dmg') or asset_name.endswith('.zip'):
                        download_url = asset.get('browser_download_url')
                        break
            
            # Compare versions
            update_available = False
            if latest_version:
                try:
                    update_available = version_parser.parse(latest_version) > version_parser.parse(self.current_version)
                except Exception:
                    # Fallback to simple string comparison if version parsing fails
                    update_available = latest_version != self.current_version
            
            return {
                'update_available': update_available,
                'latest_version': latest_version,
                'current_version': self.current_version,
                'download_url': download_url,
                'release_notes': release_notes,
                'html_url': data.get('html_url'),
                'error': None
            }
            
        except urllib.error.HTTPError as e:
            if e.code == 404:
                return {
                    'update_available': False,
                    'error': 'Repository or release not found. Please check the GitHub repository URL.'
                }
            else:
                return {
                    'update_available': False,
                    'error': f'HTTP Error {e.code}: {e.reason}'
                }
                
        except urllib.error.URLError as e:
            return {
                'update_available': False,
                'error': f'Network error: {str(e.reason)}'
            }
            
        except json.JSONDecodeError:
            return {
                'update_available': False,
                'error': 'Failed to parse GitHub response'
            }
            
        except Exception as e:
            return {
                'update_available': False,
                'error': f'Unexpected error: {str(e)}'
            }
    
    def open_download_page(self, url=None):
        """
        Open the download page in the default browser
        
        Args:
            url: Specific URL to open, or None for latest release page
        """
        if url:
            webbrowser.open(url)
        else:
            # Open the releases page
            releases_url = f"https://github.com/{self.github_repo}/releases/latest"
            webbrowser.open(releases_url)
    
    def download_and_install_update(self, download_url):
        """
        Download and begin installation of an update
        
        Args:
            download_url: URL to download the update from
            
        Returns:
            dict with keys:
                - success (bool): Whether download was initiated
                - error (str): Error message if failed
        """
        if not download_url:
            return {'success': False, 'error': 'No download URL provided'}
        
        try:
            # For macOS, open the download URL in browser
            # The user will download the .dmg and install it manually
            webbrowser.open(download_url)
            return {
                'success': True,
                'error': None,
                'message': 'Download started. Please install the update when the download completes.'
            }
            
        except Exception as e:
            return {
                'success': False,
                'error': f'Failed to open download: {str(e)}'
            }
    
    def get_current_version(self):
        """Get the current application version"""
        return self.current_version

    def _get_ssl_context(self):
        """Return an SSL context that works reliably on packaged macOS builds."""
        if certifi is not None:
            return ssl.create_default_context(cafile=certifi.where())
        return ssl.create_default_context()

    def _resolve_release_version(self, release_data, assets):
        """Resolve semantic version from GitHub release data.

        Preference order:
        1) tag_name if semver-like (e.g. v1.2.3)
        2) release name if it contains semver
        3) asset filename containing semver (e.g. App-1.2.3-macOS.dmg)
        """
        tag_name = (release_data.get('tag_name') or '').strip()
        release_name = (release_data.get('name') or '').strip()

        tag_candidate = self._extract_semver(tag_name)
        if tag_candidate:
            return tag_candidate

        name_candidate = self._extract_semver(release_name)
        if name_candidate:
            return name_candidate

        for asset in assets:
            asset_name = (asset.get('name') or '').strip()
            asset_candidate = self._extract_semver(asset_name)
            if asset_candidate:
                return asset_candidate

        # Last resort: strip leading v if present and return raw tag
        return tag_name.lstrip('v')

    @staticmethod
    def _extract_semver(text):
        """Extract X.Y.Z from text; returns None if absent."""
        if not text:
            return None
        match = re.search(r'(\d+\.\d+\.\d+)', text)
        return match.group(1) if match else None
    
    @staticmethod
    def is_running_from_bundle():
        """Check if the application is running from a macOS .app bundle"""
        return getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS')


# Convenience function for quick update checks
def check_for_updates():
    """Quick function to check for updates"""
    manager = UpdateManager()
    return manager.check_for_updates()


if __name__ == "__main__":
    # Test the update manager
    print(f"Current version: {__version__}")
    print(f"Checking for updates from: {__github_repo__}")
    
    manager = UpdateManager()
    result = manager.check_for_updates()
    
    if result.get('error'):
        print(f"\n❌ Error: {result['error']}")
    elif result.get('update_available'):
        print(f"\n✅ Update available!")
        print(f"Current version: {result['current_version']}")
        print(f"Latest version: {result['latest_version']}")
        print(f"Download URL: {result['download_url']}")
        print(f"\nRelease notes:\n{result['release_notes']}")
    else:
        print(f"\n✅ You are running the latest version ({result.get('current_version', __version__)})")
