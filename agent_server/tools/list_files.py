from agents import function_tool

from agent_server.utils import get_user_workspace_client

VOLUME_URI = "/Volumes/data-science-team-catalog/team_days/energy_reports"


def volume_file_names() -> list[str]:
    """Return the names of the files in the volume, without subfolders and .gz files."""
    # Runs on behalf of the end user, so the user needs READ VOLUME on the volume
    client = get_user_workspace_client()
    return [
        entry.name
        for entry in client.files.list_directory_contents(VOLUME_URI)
        if entry.name and not entry.is_directory and not entry.name.endswith(".gz")
    ]


@function_tool
def get_files_in_volume() -> str:
    """List the files in the energy reports volume, one file name per line."""
    names = volume_file_names()
    return "\n".join(names) if names else "The volume is empty."
