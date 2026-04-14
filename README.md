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

// From a byte slice
let bytes = std::fs::read("email.msg").unwrap();
let outlook = Outlook::from_slice(&bytes).unwrap();

// From any std::io::Read source (file, stdin, network, etc.)
let file = std::fs::File::open("email.msg").unwrap();
let outlook = Outlook::from_reader(file).unwrap();
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

### Detecting embedded messages

Attachments with `attach_method == 5` are nested `.msg` files (embedded
messages). You can identify them and handle them separately:

```rust
for attach in &outlook.attachments {
    match attach.attach_method {
        1 => println!("{}: regular file", attach.long_file_name),
        5 => println!("{}: embedded .msg", attach.display_name),
        6 => println!("{}: OLE object", attach.display_name),
        _ => println!("{}: other ({})", attach.display_name, attach.attach_method),
    }
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
