Outlook Email Message (.msg) parser
====================================
![Crates.io](https://img.shields.io/crates/v/msg_parser)
![docs.rs](https://img.shields.io/docsrs/msg_parser)

A parser for Microsoft Outlook `.msg` files (OLE Compound Document format).
Extracts message metadata, body content, recipients, attachments, and transport
headers as specified in [MS-OXMSG] and [MS-OXPROPS].

### Usage

Add this to your `Cargo.toml`:
```toml
[dependencies]
msg_parser = "0.3"
```

### Quick Start

```rust
use msg_parser::Outlook;

fn main() {
    let outlook = Outlook::from_path("email.msg").unwrap();

    println!("From: {} <{}>", outlook.sender.name, outlook.sender.email);
    println!("Subject: {}", outlook.subject);

    for person in &outlook.to {
        println!("To: {} <{}>", person.name, person.email);
    }
    for person in &outlook.cc {
        println!("CC: {} <{}>", person.name, person.email);
    }
}
```

### Parsing from different sources

```rust
use msg_parser::Outlook;

// From a file path
let outlook = Outlook::from_path("email.msg").unwrap();

// From a byte slice (accepts &[u8], Vec<u8>, or anything AsRef<[u8]>)
let bytes = std::fs::read("email.msg").unwrap();
let outlook = Outlook::from_slice(&bytes).unwrap();

// Passing a Vec<u8> directly also works
let outlook = Outlook::from_slice(bytes).unwrap();

// From any std::io::Read source (file, stdin, network, etc.)
let file = std::fs::File::open("email.msg").unwrap();
let outlook = Outlook::from_reader(file).unwrap();
```

### Display formatting

`Outlook`, `Person`, and `Attachment` all implement `Display` for
human-readable output:

```rust
let outlook = Outlook::from_path("email.msg").unwrap();
// Prints a summary: From, Subject, To, CC, BCC, Date, Attachments
println!("{}", outlook);
```

### Saving attachments

```rust
let outlook = Outlook::from_path("email.msg").unwrap();

for attach in &outlook.attachments {
    let name = if attach.long_file_name.is_empty() {
        &attach.file_name      // 8.3 short name fallback
    } else {
        &attach.long_file_name // full original filename
    };
    std::fs::write(name, &attach.payload_bytes).unwrap();
}
```

### Detecting and parsing embedded messages

Attachments with `attach_method == 5` are nested `.msg` files (embedded
messages). Use `as_message()` to parse them recursively:

```rust
let outlook = Outlook::from_path("email.msg").unwrap();

for attach in &outlook.attachments {
    if let Some(Ok(nested)) = attach.as_message() {
        println!("Embedded message subject: {}", nested.subject);
        println!("Embedded from: {}", nested.sender);
        // You can access all fields on the nested message, including
        // its own attachments (which may also be embedded .msg files)
    }
}

// Or use the convenience method:
for attach in &outlook.attachments {
    if attach.is_embedded_message() {
        println!("{} is an embedded .msg", attach.display_name);
    }
}
```

### Inline images (Content-ID)

HTML bodies reference inline images via `cid:` URIs. Use `content_id` to
resolve them:

```rust
let outlook = Outlook::from_path("email.msg").unwrap();
let mut html = outlook.html.clone();

for attach in &outlook.attachments {
    if !attach.content_id.is_empty() {
        // Replace cid: references with actual data
        let cid_ref = format!("cid:{}", attach.content_id);
        let data_uri = format!(
            "data:{};base64,{}",
            attach.mime_tag,
            base64_encode(&attach.payload_bytes), // your base64 encoder
        );
        html = html.replace(&cid_ref, &data_uri);
    }
}
```

### RTF decompression and HTML extraction

Many `.msg` files store the body as compressed RTF rather than HTML.
Use `rtf_decompressed()` to get the raw RTF, or `html_from_rtf()` to
extract embedded HTML:

```rust
let outlook = Outlook::from_path("email.msg").unwrap();

// Get the best available HTML body
let html = if !outlook.html.is_empty() {
    outlook.html.clone()
} else {
    // Many messages embed HTML inside compressed RTF
    outlook.html_from_rtf().unwrap_or_default()
};

// Or work with the raw decompressed RTF directly
if let Some(rtf_bytes) = outlook.rtf_decompressed() {
    std::fs::write("body.rtf", &rtf_bytes).unwrap();
}
```

### Message metadata

```rust
let outlook = Outlook::from_path("email.msg").unwrap();

// Timestamps (ISO 8601 UTC, empty string if unavailable)
println!("Delivered: {}", outlook.message_delivery_time);
println!("Submitted: {}", outlook.client_submit_time);
println!("Created:   {}", outlook.creation_time);
println!("Modified:  {}", outlook.last_modification_time);

// Classification
println!("Class:       {}", outlook.message_class);  // e.g. "IPM.Note"
println!("Importance:  {}", outlook.importance);      // 0=Low, 1=Normal, 2=High
println!("Sensitivity: {}", outlook.sensitivity);     // 0=Normal, 1=Personal, 2=Private, 3=Confidential
```

### JSON output

```rust
let outlook = Outlook::from_path("email.msg").unwrap();
let json = outlook.to_json().unwrap();
println!("{}", json);
```

### Available fields

| Field | Type | Description |
|-------|------|-------------|
| `headers` | `TransportHeaders` | SMTP transport headers (raw + parsed fields) |
| `sender` | `Person` | Sender name and email |
| `to` | `Vec<Person>` | Primary recipients |
| `cc` | `Vec<Person>` | Carbon-copy recipients |
| `bcc` | `Vec<Person>` | Blind carbon-copy recipients |
| `subject` | `String` | Subject line |
| `body` | `String` | Plain-text body |
| `html` | `String` | HTML body |
| `rtf_compressed` | `String` | RTF body (hex-encoded) |
| `message_class` | `String` | Message class (e.g. `"IPM.Note"`) |
| `importance` | `u32` | `0`=Low, `1`=Normal, `2`=High |
| `sensitivity` | `u32` | `0`=Normal, `1`=Personal, `2`=Private, `3`=Confidential |
| `client_submit_time` | `String` | ISO 8601 UTC timestamp |
| `message_delivery_time` | `String` | ISO 8601 UTC timestamp |
| `creation_time` | `String` | ISO 8601 UTC timestamp |
| `last_modification_time` | `String` | ISO 8601 UTC timestamp |
| `attachments` | `Vec<Attachment>` | File attachments with metadata and raw bytes |

### Attachment fields

| Field | Type | Description |
|-------|------|-------------|
| `display_name` | `String` | Display name shown in the mail client |
| `payload` | `String` | Hex-encoded attachment content |
| `payload_bytes` | `Vec<u8>` | Raw attachment bytes |
| `extension` | `String` | File extension (e.g. `".pdf"`) |
| `mime_tag` | `String` | MIME type (e.g. `"image/png"`) |
| `file_name` | `String` | Short 8.3 filename |
| `long_file_name` | `String` | Full original filename |
| `attach_method` | `u32` | `1`=file, `5`=embedded .msg, `6`=OLE object |
| `content_id` | `String` | Content-ID for inline images |

### Methods

| Method | Returns | Description |
|--------|---------|-------------|
| `Outlook::from_path(path)` | `Result<Outlook, Error>` | Parse from filesystem path |
| `Outlook::from_slice(bytes)` | `Result<Outlook, Error>` | Parse from byte slice or `Vec<u8>` |
| `Outlook::from_reader(reader)` | `Result<Outlook, Error>` | Parse from any `Read` source |
| `Outlook::to_json()` | `Result<String, Error>` | Serialize to JSON |
| `Outlook::rtf_decompressed()` | `Option<Vec<u8>>` | Decompress RTF body |
| `Outlook::html_from_rtf()` | `Option<String>` | Extract HTML from compressed RTF |
| `Attachment::as_message()` | `Option<Result<Outlook, Error>>` | Parse embedded `.msg` attachment |
| `Attachment::is_embedded_message()` | `bool` | Check if attachment is embedded `.msg` |

### Requirements

- Rust edition 2024 (rustc 1.85+)

### Running the example

```bash
cargo run --example parse-email
# or with a specific file:
cargo run --example parse-email -- path/to/email.msg
```

### Running tests

```bash
cargo test
```

### Contributions

Feel free to open pull requests to contribute, enhance, or fix bugs.

---
License: MIT
