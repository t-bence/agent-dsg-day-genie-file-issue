import base64

from agents import function_tool
from databricks_openai import AsyncDatabricksOpenAI

from agent_server.tools.list_files import VOLUME_URI
from agent_server.utils import get_user_workspace_client

IMAGE_MODEL = "databricks-claude-opus-5"
MIME_TYPES = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png"}


@function_tool
async def analyze_image(filename: str, question: str) -> str:
    """Look at an image (JPG or PNG) from the energy reports volume and answer a question about it.

    Use a file name returned by get_files_in_volume, not a full path.
    The question says what to look for, for example "Describe the chart and list its values."
    """
    if "/" in filename:
        return "Pass only the file name, without a path."
    extension = "." + filename.rsplit(".", 1)[-1].lower()
    mime_type = MIME_TYPES.get(extension)
    if mime_type is None:
        return f"Unsupported image type: {filename}. Supported types: {', '.join(MIME_TYPES)}."

    # Runs on behalf of the end user, so the user needs READ VOLUME on the volume
    client = get_user_workspace_client()
    response = client.files.download(f"{VOLUME_URI}/{filename}")
    if response.contents is None:
        return f"The file {filename} is empty."
    data_url = f"data:{mime_type};base64,{base64.b64encode(response.contents.read()).decode()}"

    # The Agents SDK drops images from tool outputs on the chat completions API,
    # so the tool sends the image to the model itself and returns the text answer
    completion = await AsyncDatabricksOpenAI().chat.completions.create(
        model=IMAGE_MODEL,
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": question},
                    {"type": "image_url", "image_url": {"url": data_url}},
                ],
            }
        ],
        extra_body={"thinking": {"type": "disabled"}},
    )
    return completion.choices[0].message.content or "The model returned no answer."
