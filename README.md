Outlook Email Messsage (.msg) parser.
=======================================
![Crates.io](https://img.shields.io/crates/v/msg_parser)
![docs.rs](https://img.shields.io/docsrs/msg_parser)
[![Build Status](https://travis-ci.com/marirs/msg-parser-rs.svg?branch=master)](https://travis-ci.com/marirs/msg-parser-rs)

A simple parser and reader to deserialize a given Outlook Email Message (.msg) File.  

### Usage
Add this to your `Cargo.toml` file:
```toml
[dependencies]
msg_parser = "0.1.0"
```

### Example

```rust
use msg_parser::Outlook;

fn main() {
    // Create Outlook object
    let outlook = Outlook::from_path("data/test_email.msg").unwrap();

    // Flush as json string
    let json_string = outlook.to_json();

    println!("{:#?}", outlook);
    
    println!();
    println!("json_string ---");
    println!("{:?}", json_string);
}

```

### Requirements
- Rust 1.42+

### Running the given example
```bash
$ cargo run --example parse-email
   Compiling msg_parser v0.1.0 (msg-parser)
    Finished dev [optimized + debuginfo] target(s) in 3.66s
     Running `target/debug/examples/parse-email`
Outlook {
    headers: TransportHeaders {
        content_type: "",
        date: "",
        message_id: "",
        reply_to: "",
    },
    sender: Person {
        name: "",
        email: "",
    },
    ...<clip>
}
```

### Running tests
```bash
cargo t --verbose
```

### Building release
```bash
cargo b --release
```

### Contributions

Feel free to make pull requests to contribute/enhance/add more features/bug fixes.

---
License: MIT