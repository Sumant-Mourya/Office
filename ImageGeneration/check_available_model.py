from google import genai
import config

client = genai.Client(api_key=config.GEMINI_API_KEY)

def is_image_model(model):
    name = model.name.lower()

    # Strong detection rules
    image_keywords = [
        "imagen",
        "image",
        "generate",
    ]

    # Explicit include patterns (most reliable)
    if "imagen" in name:
        return True

    if "image" in name:
        return True

    if "generate" in name and (
        "imagen" in name or "veo" in name
    ):
        return True

    return False


def main():
    print("\n🔍 Fetching available models...\n")

    models = client.models.list()

    all_models = []
    image_models = []

    for m in models:
        name = m.name
        all_models.append(name)

        if is_image_model(m):
            image_models.append(name)

    # Print all models
    print("=" * 60)
    print("📦 ALL AVAILABLE MODELS")
    print("=" * 60)
    for m in sorted(all_models):
        print(m)

    # Print detected image models
    print("\n" + "=" * 60)
    print("🖼️ IMAGE GENERATION MODELS")
    print("=" * 60)

    if image_models:
        for m in sorted(image_models):
            print(m)
    else:
        print("❌ No image generation models detected")


if __name__ == "__main__":
    main()